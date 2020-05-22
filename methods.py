import pandas as pd
import pyodbc
import itertools
import schedule
import time
import ctypes
import glob
import os
import smtplib
import matplotlib.pyplot as plt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import win32com.client as win32
from wordcloud import WordCloud, STOPWORDS
import matplotlib.gridspec as gridspec
import numpy as np
import win32com.client
import re
import logging


class Methods:
    def __init__(self):
        self.smtp_host = 'smtp-goss.int.worlf.socgen'
        self.smtp_port = 25
        self.smtp_timeout = 10

# ----------------------------------------------------PANDAS-----------------------------------------------------------
    @staticmethod
    def file_to_df(file_path):
        """Export File Content into DataFrame"""
        my_list = []
        file_name, file_ext = os.path.splitext(os.path.basename(file_path))
        if file_ext in ['.xls', '.xlsx', '.xlsm']:
            my_list = pd.read_excel(file_name)
        elif file_ext in ['.csv']:
            my_list = pd.read_csv(file_name)
        elif file_ext in ['.txt']:
            my_list = pd.read_csv(file_name, header=None, na_values='', index_col=None, sep='\n')
            my_list = [line.split(';') for line in my_list[0] if str(line) != 'nan']
            my_list = pd.DataFrame(my_list[1:], columns=[elem.replace('"', '') for elem in my_list[0]])

        return my_list

    @staticmethod
    def query_to_df(query, server_name, db_name):
        """Export SQL Query data into DataFrame"""
        conn = pyodbc.connect(DRIVER='{SQL Server}', SERVER=server_name, DATABASE=db_name)
        cursor = conn.cursor()
        cursor.execute(query)
        column_names = [col[0] for col in cursor.description]
        my_list = pd.DataFrame([dict(itertools.zip_longest(column_names, row)) for row in cursor.fetchall()])
        return my_list

    @staticmethod
    def df_comparison(df1, df2, common_col):
        """ Comparison between 2 DataFrames"""
        return df1.assign(InDf2=df1[common_col].isin(df2[common_col]).assign(int))

# ----------------------------------------------------MAIL-----------------------------------------------------------
    @staticmethod
    def recipient_list_in_txt_file(file_name):
        """Convert txt files content into List of Recipients"""
        rec_file = glob.glob(file_name)
        rec_file = [line.strip() for line in open(rec_file[0], 'r')]
        return rec_file

    @staticmethod
    def mail_with_win32(to_mail, sub_mail, html):
        """ Send Mail via Win32"""
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.To = to_mail
        mail.Subject = sub_mail
        mail.HTMLBody = html
        mail.Send()

    def mail_with_picture(self, to_mail, from_mail, sub_mail, chart_path, html):
        """Send Mail with picture via SMTP"""
        msg = MIMEMultipart('related')
        msg['Subject'] = sub_mail
        part_html = MIMEText(html, 'html')
        msg.attach(part_html)

        fp = open(chart_path, 'rb')
        msg_image = MIMEImage(fp.read())
        fp.close()
        msg_image.add_header('Content-ID', '<image1>')
        msg.attach(msg_image)

        smt = smtplib.SMTP(host=self.smtp_host, port=self.smtp_port, timeout=self.smtp_timeout)
        smt.sendmail(from_mail, to_mail, msg.as_string())
        smt.quit()

    @staticmethod
    def import_attachment_mail(extension_list, subject_mail, path_to_paste_attachment):
        """ Import all the attachments received per Mail"""
        outlook = win32com.client.Dispatch('Outlook.Application').GetNamespace('MAPI')
        inbox = outlook.GetDefaultFolder(6)
        # All the attachments already imported in the destination folder
        extract_imported = [os.path.basename(x) for x in glob.glob(path_to_paste_attachment + '\\*')]

        for message in inbox.Items:
            if subject_mail in message.Subject:
                for attachment in message.Attachments:
                    file_name, ext = os.path.splitext(str(attachment))
                    if ext in extension_list and str(attachment) not in extract_imported:
                        attachment.SaveAsFile(os.path.join(path_to_paste_attachment, str(attachment)))
                        extract_imported = [os.path.basename(x) for x in glob.glob(path_to_paste_attachment + '\\*')]

# ----------------------------------------------------FORMAT-----------------------------------------------------------
    @staticmethod
    def format_tab_into_float(df, columns_to_exclude):
        """Convert String into float with separator"""
        for col in df.columns:
            if col not in columns_to_exclude:
                df[col] = ['{:,}'.format(round(float(u[col]))) for i, u in df.iterrows()]

    @staticmethod
    def format_tab_mail(df, format_mail_color, format_col):
        """Format DataFrame into Mail Body"""
        return (df.style
                .set_table_attribute('class="table"')
                .apply(format_mail_color, col=format_col, subset=format_col)
                .set_properties(**{'text-align': 'center', 'border-color': 'Black', 'border-width': 'thin',
                                   'border-style': 'dotted', 'width': '130px'})
                .set_table_styles([{'selector': 'th',
                                    'props': [('font-size', '12pt'), ('border-style', 'dotted'),
                                              ('background-color', 'whitesmoke'), ('border-width', '1px')]}])
                .hide_index()
                .render())

    @staticmethod
    def format_mail_color(val, col):
        """Color Code into Mail Body"""
        return ['background-color: tomato' if v <= 0 and val.name == col
                else 'background-color: lightgreen' for v in val]

    @staticmethod
    def format_df_coma(df, column_name):
        """Format a DataFrame column to #,###.00"""
        df[column_name] = df[column_name].apply(lambda x: '{:,.2f}'.format(float(x)))

# ------------------------------------------------FILES/FOLDER----------------------------------------------------------
    @staticmethod
    def concat_files_in_folder(path_files, pattern):
        """ Concatenation of all the files in a specific folder """
        # pattern = 'CASH_BALANCE_(.*)_REPORT_(.*)'
        extract_imported = glob.glob(path_files + '\\*')
        df = pd.DataFrame([])
        for file in extract_imported:
            file_name, ext = os.path.splitext(os.path.basename(file))
            pattern = re.search(pattern, file_name.upper())
            content_file = pd.read_csv(file) if ext == '.csv' else pd.read_excel(file)
            df = df.append(content_file)

    @staticmethod
    def check_folder_content(folder_path, error_msg):
        """Generate a MsgBox if folder empty"""
        filename = glob.glob(folder_path + '\\*')
        if filename:
            ctypes.windll.user32.MessageBoxW(0, error_msg, "Error", 1)
            raise SystemExit()

    @staticmethod
    def move_files(initial_path, destination_path):
        """Move Files into another Folder"""
        os.rename(initial_path, destination_path)

    @staticmethod
    def remove_files_in_folder(path):
        """Remove all the files in Folder"""
        files_name = [glob.glob(folder + '//*') for folder in glob.glob(path)]
        for file_path in files_name:
            for sub_file_path in file_path:
                path, ext = os.path.splitext(sub_file_path)
                os.remove(sub_file_path)

# ------------------------------------------------CHART-----------------------------------------------------------------
    @staticmethod
    def chart_word_cloud(list_words, chart_path, chart_title):
        """WordCLoud Chart based on DF"""
        stopwords = set(STOPWORDS)
        word_cloud = WordCloud(background_color='white', colormap='Blues', stopwords=stopwords)\
            .generate(list(list_words))

        plt.figure(figsize=(12, 10))
        plt.title(chart_title, fontsize=20)
        plt.axis('off')
        plt.imshow(word_cloud, interpolation='bilinear')
        plt.savefig(chart_path)

    @staticmethod
    def chart_multi_picture():
        """ Multi Pictures"""
        fig = plt.figure(figsize=(12, 9), edgecolor="grey", linewidth=2)
        gs = gridspec.GridSpec(3, 2)
        fig.add_subplot(gs[0, :])
        fig.add_subplot(gs[1, 0])
        fig.add_subplot(gs[1, 1])

    @staticmethod
    def chart_multi_curves(axis_x, df, chart_path, columns):
        fig = plt.figure(figsize=(12, 9), edgecolor="grey", linewidth=2)
        x = axis_x
        plt.style.use('seaborn-darkgrid')
        plt.plot(x, df[columns].iloc[::-1], 'b:o', linewidth=3, alpha=0.6, label='exemple')
        plt.plot(x, [0 for w in x], 'k-', linewidth=1, alpha=0.6)
        plt.fill_between(x, df[columns].iloc[::-1], [0 for w in x], alpha=0.2, facecolor='blue')

        plt.xlabel("Axis X")
        plt.ylabel("Axis Y")
        plt.title("Title")

        plt.gcf().autofmt_xdate()  # Axis X Adjustment
        plt.xticks(x[::3], x[::3])

        plt.grid(alpha=0.2, linestyle='--', color='grey')
        plt.legend()
        fig.savefig(chart_path)

# ------------------------------------------------SCHEDULER/ LOG--------------------------------------------------------
    @staticmethod
    def scheduler_parameters(func, time_interval):
        schedule.every(time_interval).minutes.do(func)
        while True:
            schedule.run_pending()
            time.sleep(time_interval)

    @staticmethod
    def log_function(log_path):
        """Read and Write into Log File"""
        logging.basicConfig(filename=log_path, level=logging.DEBUG)
        with open(log_path) as log_file:
            log_data = log_file.readlines()

        logging.info('BREACH')


# ------------------------------------------------SQL------------------------------------------------------------------

sql_remove_duplicates = """WITH CTE AS ( 
                SELECT [RunDate], [Deal Number], 
                RN = ROW_NUMBER() OVER (PARTITION BY [RunDate], [Deal Number] ORDER BY [RunDate], [Deal Number])
                FROM TableA)
                DELETE FROM CTE WHERE RN>1"""

# ------------------------------------------------GitHub------------------------------------------------------------------
"""
version:
    git --version
configuration initial:
    git config --global user.name "joanserfaty"
    git config --global user.email "joan.serfaty@hotmail.com"
create a project:
    mkdir PythonProjects
select the project:
    cd PythonProjects
Initialiser git:
    git init
Add Content to the Git repository:
    git add methods.py
Save our amendments:
    git commit -m "project initi without"
Check all the historical amendments:
    git log
Create branch:
    git branch with_git_comments
    git branch
Change branch + Amendments:
    git checkout with_git_comment
    git add methods.py
    git commit -m "new version"
Merge to the master branch:
    git checkout master
    git merge with_git_comment
    git branch -d with_git_comment
Create a repertory based on the existing repertory on Github:
    git remote add origin https://github.com/joanserfaty/python_recap.git
    git push -u origin master
Create a new repertory on GitHub:
    echo "# python_recap" >> README.md
    git init
    git add README.md
    git commit -m "first commit"
    git remote add origin https://github.com/joanserfaty/python_recap.git
    git push -u origin master
Bring back the project of a colleague:
    git pull PythonProjects master
    

"""
