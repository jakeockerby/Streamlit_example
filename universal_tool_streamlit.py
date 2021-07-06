# -*- coding: utf-8 -*-
"""
Created on 29/06/2021
universal_tool_streamlit.py

The purpose of this program is to provide a universal interactive tool that 
automates the process of retrieving values from the portal and give further
insight into the data.
 
General age, gender, product and date filters are provided as well as a
question time type feature and graph filters.

@author: Jake Ockerby 
The Insights People

===========================Version Control====================================
Number=========================Author=======================Notes=============
V2                              JAO                    Initial Release

"""


import seaborn as sns
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import configparser
import pyodbc
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from openpyxl.styles import Font
from time import sleep
import datetime
from fuzzywuzzy import process
from scipy.stats.mstats import winsorize
import datetime
from dateutil.relativedelta import relativedelta
import itertools
import string
import os
import pandas as pd
import numpy as np
import streamlit as st
from bokeh.plotting import figure
import matplotlib.pyplot as plt
import datetime
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go



# Initialize the ConfigParser
config = configparser.ConfigParser()
# Read the config
#config.read('/home/portaladmin/scripts/python/Data_Science/config')
config.read(r'Z:/Technology/Data Science/Config/config.txt')
username = config.get('DEFAULT', 'username')
password = config.get('DEFAULT', 'password')
server = config.get('DEFAULT', 'server')
database = config.get('DEFAULT', 'database')
# Create connections to the SQL db
# Adjust the parameters to connect to different SQL dbs
connection_string = 'DRIVER={ODBC Driver 17 for SQL Server};'
connection_string += 'SERVER=' + server + ';'
connection_string += 'DATABASE=' + database + ';'
connection_string += 'UID=' + username + ';'
connection_string += 'PWD=' + password + ';'
cnxn = pyodbc.connect(connection_string)
cursor = cnxn.cursor()



# Function to get the hours and minutes from a datetime and putting it in a
# suitable format
def hour_mins(x):
    hours = x.split(':')[0]
    mins = x.split(':')[1]
    if hours == "'0":
        new_format = '{}mins'.format(mins)
    else:
        new_format = '{0}h {1}mins'.format(hours, mins)
    
    return new_format


# Function for MONY011 question code
def money011(old_num, new_num):
    old_num['answer'] = old_num['answer'].astype(float)
    new_num['answer'] = new_num['answer'].astype(float)
 # List to store winsorized answers
    winsorized_old = []
    
    # For each subquestion, winsorize the answers with limits set to 2 s.d
    # above and below the mean and then append the results to the list
    for i in list(old_num['subquestion'].unique()):
        winsor = old_num.loc[old_num['subquestion'] == i]
        winsor['answer'] = winsorize(winsor['answer'], limits=[0.025, 0.025])
        winsorized_old.append(winsor)
    
    # Concatenate the results for each subquestion
    winsorized_old = pd.concat(winsorized_old, ignore_index=True)
    
    # Group by subquestion, retrieve the answer column and calculate the
    # mean for each subquestion
    before = winsorized_old.groupby(['subquestion'])['answer'].mean()
 
    
    # Same as above, but applied to new values
    winsorized_new = []
    for i in list(new_num['subquestion'].unique()):
        winsor = new_num.loc[new_num['subquestion'] == i]
        winsor['answer'] = winsorize(winsor['answer'], limits=[0.025, 0.025])
        winsorized_new.append(winsor)
        
    winsorized_new = pd.concat(winsorized_new, ignore_index=True)
    
    
    after = winsorized_new.groupby(['subquestion'])['answer'].mean()
    
    # Making dataframe to display values for each time period
    before = pd.DataFrame({'name':before.index, 'value':before.values})
    after = pd.DataFrame({'name':after.index, 'value':after.values})
    
    # Mergeing the dataframes ready for comparison
    df_merge_col = pd.merge(before, after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
    df_merge_col = df_merge_col.loc[df_merge_col['before'] > 0]
    # Calculating growth
    df_merge_col['growth'] = df_merge_col.apply(lambda x: (x['after'] / 
                                                      x['before'])-1, axis=1)
            
    # Scaling by 100          
    df_merge_col['growth'] = df_merge_col['growth'].apply(lambda x: x*100)
    return df_merge_col


# Function for TIME005 question code
def time005(old_num, new_num):
    # Calculating mean for each subquestion
    old_num['answer'] = old_num['answer'].astype(float)
    new_num['answer'] = new_num['answer'].astype(float)
    
    before = old_num.groupby(['subquestion'])['answer'].mean()
    after = new_num.groupby(['subquestion'])['answer'].mean()
    
    # Making dataframe to display values for each time period
    before = pd.DataFrame({'name':before.index, 'value':before.values})
    after = pd.DataFrame({'name':after.index, 'value':after.values})
    
    # Mergeing the dataframes 
    df_merge_col = pd.merge(before, after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
    df_merge_col = df_merge_col.loc[df_merge_col['before'] > 0]
    # Calculating growth
    df_merge_col['growth'] = df_merge_col.apply(lambda x: (x['after'] / 
                                                      x['before'])-1, axis=1)
               
    df_merge_col['growth'] = df_merge_col['growth'].apply(lambda x: x*100)
    return df_merge_col


# Function for FVRT006 question code
def fvrt006(old_num, new_num, qcode):
    
    # qcode = qcode[0]
    # Selecting entries either on PC/Console or Mobile and swapping the answer
    # and subquestion columns
    if qcode == 'FVRT006A':
        old_num = old_num.loc[old_num['answer'] == 'PC and Console']
        old_num = old_num.rename(columns={'answer': 'change1', 'subquestion': 'change2'})
        old_num = old_num.rename(columns={'change1': 'subquestion', 'change2': 'answer'})
        
        new_num = new_num.loc[new_num['answer'] == 'PC and Console']
        new_num = new_num.rename(columns={'answer': 'change1', 'subquestion': 'change2'})
        new_num = new_num.rename(columns={'change1': 'subquestion', 'change2': 'answer'})

    else:
        old_num = old_num.loc[old_num['answer'] == 'Mobile (app)']
        old_num = old_num.rename(columns={'answer': 'change1', 'subquestion': 'change2'})
        old_num = old_num.rename(columns={'change1': 'subquestion', 'change2': 'answer'})
        
        new_num = new_num.loc[new_num['answer'] == 'Mobile (app)']
        new_num = new_num.rename(columns={'answer': 'change1', 'subquestion': 'change2'})
        new_num = new_num.rename(columns={'change1': 'subquestion', 'change2': 'answer'})

    return old_num, new_num



# Function to select subquestion
def subquestions(old_num, new_num, subq):
    # Correcting misspellings
    sub_list = list(old_num['subquestion'].unique())
    new_subq = process.extractOne(subq, sub_list)[0]
    
    # Locating on the subquestion
    old_num = old_num.loc[old_num['subquestion'] == new_subq]
    new_num = new_num.loc[new_num['subquestion'] == new_subq]
    return old_num, new_num



# Function to add halos to the bars in bar graph
def bar_labels(ax, time='n', orientation='vertical', space=0.4, fontsize=12):
    # If bar graph is vertical
    if orientation == 'vertical':
        for p in ax.patches:
            x = p.get_x() + p.get_width() / 2 # Plotting at centre of bar
            y = p.get_y() + p.get_height() + float(space) # Plotting halo just above bar
            
            # If TIME005 is selected convert height into hours and mins
            if time == 'y':
                # If there is a height, convert to datetime, else set to zero
                if np.isnan(p.get_height()) == False:
                    value = str(datetime.timedelta(hours=p.get_height()))
                else:
                    value = str(datetime.timedelta(hours=0))
                
                # Pass value into above hour_mins function, which retrieves
                # hours and mins from the datetime
                value = hour_mins(value)
            else:
                # Else round value to 1 decimal place
                value = float(round(p.get_height(), 1))
            # Align at center with the fontsize passed
            ax.text(x, y, value, ha="center", fontsize=fontsize)
    else:
        # If graph is horizontal
        for p in ax.patches:
            x = p.get_x() + p.get_width() + float(space) # Plotting halo just above bar
            y = p.get_y() + p.get_height() / 2 # Plotting at centre of bar
            
            # If TIME005 is selected convert height into hours and mins
            if time == 'y':
                # If there is a width, convert to datetime, else set to zero
                if np.isnan(p.get_width()) == False:
                    value = str(datetime.timedelta(hours=p.get_width()))
                else:
                    value = str(datetime.timedelta(hours=0))
                
                # Pass value into above hour_mins function, which retrieves
                # hours and mins from the datetime
                value = hour_mins(value)
            else:
                # Else round value to 1 decimal place
                value = float(round(p.get_width(), 1))
            
            # If value is < 0 (for growth) align to right, else align to left
            if value < 0:
                ax.text(x, y, value, ha="right", fontsize=fontsize)
            else:
                ax.text(x, y, value, ha="left", fontsize=fontsize)
            
            

            
# Function to manipulate table based on filters the user has selected
def table_type(df, selnames, top, bottom, threshold, sign, sort):
    # Selecting entries to display by names the user has inputted
    if selnames != 0:
        return st.write('Passed loop')
        # Spellchecker
        names = list(df['name'].unique())
        selnames = selnames.split(',')
        
        new_selnames = []
        for n in selnames:
            n = process.extractOne(n, names)[0]
            new_selnames.append(n)
            
        final = df.loc[df['name'].isin(new_selnames)]
    
    # Selecting top answers
    elif top != None:
        final = df.head(int(top))
    
    # Selecting bottom answers
    elif bottom != None:
        final = df.head(int(bottom))
    
    
    # Selecting threshold
    elif threshold != None:
        if sort != 'name':
            if sign == '<':
                final = df.loc[df[sort] < threshold/100]
            elif sign == '>':
                final = df.loc[df[sort] > threshold/100]
            else:
                final = df.loc[df[sort] == threshold/100]
        else:
            final = df
            
    else:
        final = df
        
    return final



# Defining a month function that replaces month numbers with month names
def month_conv(x, by_quarter):
    if by_quarter == 'No':
        xsplit = x.split('/')
        month_dict = {'01': 'Jan',
                      '02': 'Feb',
                      '03': 'Mar',
                      '04': 'Apr',
                      '05': 'May',
                      '06': 'Jun',
                      '07': 'Jul',
                      '08': 'Aug',
                      '09': 'Sep',
                      '10': 'Oct',
                      '11': 'Nov',
                      '12': 'Dec'}
        
        new_x = x.replace(xsplit[0], month_dict[xsplit[0]], 1)
    else:
        xsplit = x.split('/')
        month_dict = {'01': 'Q1',
                      '04': 'Q2',
                      '07': 'Q3',
                      '10': 'Q4',}
        
        xsplit = xsplit
        newer_x = x.replace(xsplit[0], month_dict[xsplit[0]], 1)
        new_xsplit = newer_x.split('/')
        new_x = new_xsplit[1] + '/' + new_xsplit[0]
        # print(new_x)
        # sleep(2)
    return new_x



# Function to get last day of the month for line graph
def last_day_of_month(any_day, to_string=True):
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
    if to_string == True:
        return (next_month - datetime.timedelta(days=next_month.day)).strftime('%Y/%m/%d')
    else:
        return (next_month - datetime.timedelta(days=next_month.day))

# Function to create line graph
def line_graph(by_quarter, smooth, filepath, newdate2, num, denom, month_range, product, sumavg, selnames, text_colour, graph_title, title_font, title_fontsize, xaxis, yaxis, label_font, label_fontsize, tick_rotation, ticksize):
    if by_quarter == 'No':
        # Converting date string to datetime 
        thismonth = datetime.datetime.strptime(newdate2, '%Y/%m/%d').date()
        thismonth = thismonth.replace(day=1)
        
        # Creating lists to store nums, denoms and date ranges
        all_month_data_nums = []
        all_month_data_denoms = []
        dateranges = []
        
        # For each month in the range the user has input, get month and append to 
        # dateranges, collect the num and denom between the first and last days of 
        # the month and append to the corresponding lists
        for month in range(0, month_range + 1):
            daterange = thismonth + relativedelta(months=-month)
            daterange_str = daterange.strftime('%Y/%m/%d')
            dateranges.append(daterange)
            month_data_num = num[daterange_str:last_day_of_month(daterange)]
            month_data_denom = denom[daterange_str:last_day_of_month(daterange)]
            all_month_data_nums.append(month_data_num)
            all_month_data_denoms.append(month_data_denom)

    else:
        thismonth = datetime.datetime.strptime(newdate2, '%Y/%m/%d').date()
        thismonth = thismonth.replace(day=1) + relativedelta(months=-((thismonth.month%3) + 2))
        
        # Creating lists to store nums, denoms and date ranges
        all_month_data_nums = []
        all_month_data_denoms = []
        dateranges = []
        
        # For each month in the range the user has input, get month and append to 
        # dateranges, collect the num and denom between the first and last days of 
        # the month and append to the corresponding lists
        for month in range(0, month_range + 1, 3):
            daterange = thismonth + relativedelta(months=-month)
            daterange_str = daterange.strftime('%Y/%m/%d')
            dateranges.append(daterange)
            month_data_num = num[daterange_str:last_day_of_month(daterange + relativedelta(months=2))]
            month_data_denom = denom[daterange_str:last_day_of_month(daterange + relativedelta(months=2))]
            all_month_data_nums.append(month_data_num)
            all_month_data_denoms.append(month_data_denom)

        
    # List to store the answer dataframes
    answer_dfs = []
    
    # For all dateranges, nums and denoms calculate percentages and append to list
    for daterange, newnum, newdenom in zip(dateranges, all_month_data_nums, all_month_data_denoms):
        percs = 100*(newnum['answer'].value_counts() / len(newdenom))
        percs = pd.DataFrame({'month': daterange.strftime('%m/%Y'), 'name':percs.index, 'value':percs.values})
        
        # If quarter is set to yes, pass 'month' to month_conv function with by_quarter = 'Yes'
        if by_quarter == 'Yes':    
            percs['month'] = percs['month'].apply(lambda x: month_conv(x, by_quarter))
        
        # Append percs to dataframe list
        answer_dfs.append(percs)
    
    
    # Concatenating the dataframes into one dataframe, form a pivot table and
    # sort them into date order
    answer_table = pd.concat(answer_dfs[::-1], ignore_index=True)
    
    # if qurater is set to no, add first day of the month to date
    if by_quarter == 'No':
        answer_table['month'] = answer_table['month'].apply(lambda x: '1/' + x)
    
    
    # Filling in missing names with 0 so it doesn't mess up the data when cubic 
    # interpolation is applied
    for m in list(answer_table['month'].unique()):
        name_per_month = list(answer_table['name'].loc[answer_table['month'] == m].unique())
        names = list(answer_table['name'].unique())
        missing_names = list(set(names) - set(name_per_month))
        for n in missing_names:
            new_name = pd.DataFrame(data={'month': [m], 'name': [n], 'value': [0]})
            answer_table = answer_table.append(new_name, ignore_index=True)
    
    # If quarter and smooth is set to yes, add a letter starting at 'A' to each date,
    # making new dates ready to perform the cubic interpolation
    if by_quarter == 'Yes':
        if smooth == 'Yes':
            for m in list(answer_table['month'].unique())[:-1]:
                for n in list(answer_table['name'].unique()):
                    for letter in string.ascii_uppercase:
                        d_list = []
                        d_list.append(m + '/{}'.format(letter))
                        d_list.append(n)
                        d_list.append(np.nan)
                        new_dates = pd.DataFrame([d_list], columns=list(answer_table.columns))
                        answer_table = answer_table.append(new_dates)
    else:
        # If quarter is no and smooth is yes, add the range of days that exist in that month,
        # so that cubic interpolation can be performed
        if by_quarter == 'No':
            if smooth == 'Yes':
                for m in list(answer_table['month'].unique())[:-1]:
                    for n in list(answer_table['name'].unique()):
                        if m.split('/')[1] == '02':
                            for no in range(2, 29):
                                d_list = []
                                d_list.append(m.replace(m.split('/')[0], '0{}'.format(no), 1) if no < 10 else m.replace(m.split('/')[0], '{}'.format(no), 1))
                                d_list.append(n)
                                d_list.append(np.nan)
                                new_dates = pd.DataFrame([d_list], columns=list(answer_table.columns))
                                answer_table = answer_table.append(new_dates)
                        elif m.split('/')[1] in ['04', '06', '09', '11']:
                            for no in range(2, 31):
                                d_list = []
                                d_list.append(m.replace(m.split('/')[0], '0{}'.format(no), 1) if no < 10 else m.replace(m.split('/')[0], '{}'.format(no), 1))
                                d_list.append(n)
                                d_list.append(np.nan)
                                new_dates = pd.DataFrame([d_list], columns=list(answer_table.columns))
                                answer_table = answer_table.append(new_dates)
                        else:
                            for no in range(2, 32):
                                d_list = []
                                d_list.append(m.replace(m.split('/')[0], '0{}'.format(no), 1) if no < 10 else m.replace(m.split('/')[0], '{}'.format(no), 1))
                                d_list.append(n)
                                d_list.append(np.nan)
                                new_dates = pd.DataFrame([d_list], columns=list(answer_table.columns))
                                answer_table = answer_table.append(new_dates)                        
   
    
   # Pivot the answer table by month and name
    pivot_answers = answer_table.pivot('month', 'name', 'value')
    pivot_answers['date'] = pivot_answers.index
    
    # Sort by date and reset index
    pivot_answers = pivot_answers.sort_values(by='date')
    pivot_answers = pivot_answers.drop(['date'], axis=1)    
    pivot_answers.reset_index(inplace=True)
    
    # If smooth is yes, make a new column that is filled with integers and set as
    # index so that data is arranged in the correct order in the graphs
    if smooth == 'Yes':
        pivot_answers['no.'] = [n for n in range(len(pivot_answers))]
        pivot_answers = pivot_answers.set_index('no.')
    
    # Rearrange the dates and perform cubic interpolation
    if by_quarter == 'No':
        pivot_answers['month'] = pd.to_datetime(pivot_answers['month'], format='%d/%m/%Y')
        pivot_answers = pivot_answers.sort_values(by='month')
        if smooth == 'Yes':
            pivot_answers = pivot_answers.set_index('month')
            pivot_answers = pivot_answers.astype(float)
            pivot_answers = pivot_answers.interpolate(method='cubic')
    else:
        # Replace '/' with '-' and perform cubic interpolation
        pivot_answers['month'] = pivot_answers['month'].apply(lambda x: x.replace('/', '-'))
        if smooth == 'Yes':
            pivot_answers['no.'] = [n for n in range(len(pivot_answers))]
            pivot_answers = pivot_answers.set_index('no.')
            pivot_answers = pivot_answers.interpolate(method='cubic')
    
    # If quarter is no, rearrange the dates into a suitable format ready to display on graph
    if by_quarter == 'No':
        pivot_answers.reset_index(inplace=True)    
        pivot_answers['month'] = pivot_answers['month'].apply(lambda x: x.strftime('%m/%Y/%d'))
        pivot_answers['month'] = pivot_answers['month'].apply(lambda x: month_conv(x, by_quarter='No'))
        pivot_answers['month'] = pivot_answers['month'].apply(lambda x: x.split('/')[0] + '-' + x.split('/')[1] if x.split('/')[2] == '01' else x)
    
    pivot_answers = pivot_answers.set_index('month')
        
    if smooth == 'No':
        # Fill NaN values with 0 so graph plots correctly
        pivot_answers = pivot_answers.fillna(0)
    
    # Spellchecker
    names = list(pivot_answers.columns)
    selnames = selnames.split(',')
    
    new_selnames = []
    for n in selnames:
        n = process.extractOne(n, names)[0]
        new_selnames.append(n)
    
    pivot_answers = pivot_answers[new_selnames]

    # Colour wheel
    colors = ['#96D5EE', '#27AAE1', '#2D2E83', '#7277C7',
              '#00A19A', '#ACCC00', '#662483', '#9D96C9',
              '#CE2F6C', '#FAA932', '#464B9A', '#347575',
              '#6AC2C2', '#76AF41', '#A33080', '#EB8EDB',
              '#F66D9B', '#D95D27', '#F7915D', '#FFC814'] 
    
    
    # Setting user inputted text colour
    plt.rcParams['text.color'] = text_colour
    
    plt.figure(figsize=(14, 6))
    graph = sns.lineplot(data=pivot_answers, dashes=False, sort=False,
                         legend='full',
                         palette=colors[:len(new_selnames)])
    
    # Fonts for title and labels
    tfont = {'fontname':'{}'.format(title_font)}
    lfont = {'fontname':'{}'.format(label_font)}
    
    # Set user inputted title with correct formating
    plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
    
    
    # Adding correct formating to graph labels
    plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
    plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
    
    # Rotate x-axis ticks based on what user has
    # selected for tick_rotation
    plt.xticks(rotation=tick_rotation)
    
    ax = plt.gca()
    
    # Setting colour of the tick labels       
    [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
    [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
    [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
    [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
    
    ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
    
    # Only showing ticks we want
    if by_quarter == 'No':
        if smooth == 'Yes':
            pivot_answers.reset_index(inplace=True)
            new_ticks = []
            count = 3
            for m in pivot_answers['month']:
                if len(m.split('/')) > 2:
                    m = ''
                if len(m.split('-')) == 2:
                    count += 1
                    if count != 4:
                        m = ''
                    else:
                        m = m
                        count = 0
                new_ticks.append(m)

            plt.xticks(pivot_answers['month'], new_ticks)
        else:
            n = 4
            [l.set_visible(False) for (i,l) in enumerate(ax.xaxis.get_ticklabels()) if i % n != 0]
    else:
        if smooth == 'Yes':
            n = 27
            [l.set_visible(False) for (i,l) in enumerate(ax.xaxis.get_ticklabels()) if i % n != 0]

     
        
    plt.tick_params(axis = "x", which = "both", bottom = False, top = False)
    
    # Adding legend
    plt.legend(bbox_to_anchor=(0.6, -0.1), ncol=len(new_selnames), frameon=False).set_title('')
    plt.xticks(fontsize=ticksize)
    plt.yticks(fontsize=ticksize)
    
    # Despine graphs
    sns.despine(left=True)
    
    # Getting current time for filenames
    current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
    
    # Save graphs
    if sumavg == 'Average':
        graph.figure.savefig('{0}line_graph_avg_{1}.png'.format(filepath, current_time), bbox_inches='tight', transparent=True)
    else:
        graph.figure.savefig('{0}line_graph_{1}_{2}.png'.format(filepath, product, current_time), bbox_inches='tight', transparent=True)
    



def pie_chart(filepath, final, sumavg, product, text_colour, graph_title, title_fontsize, title_font, pie_label_col):
    # Colour wheel
    colors = ['#96D5EE', '#27AAE1', '#2D2E83', '#7277C7',
              '#00A19A', '#ACCC00', '#662483', '#9D96C9',
              '#CE2F6C', '#FAA932', '#464B9A', '#347575',
              '#6AC2C2', '#76AF41', '#A33080', '#EB8EDB',
              '#F66D9B', '#D95D27', '#F7915D', '#FFC814']
    
    for col in ['before', 'after']:
        # Setting user inputted text colour
        plt.rcParams['text.color'] = text_colour
        plt.figure(figsize=(20,6))
        
        # Fonts for title and labels
        tfont = {'fontname':'{}'.format(title_font)}
        
        # Set user inputted title with correct formating
        plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize, y=1.07)
        
        pie, q, w = plt.pie(final[col], autopct='%1.0f%%', pctdistance=0.75,
                labeldistance=1.35, colors=colors, counterclock=False, startangle=90,
                explode=([0.01]*len(final['name'].unique())), textprops={'color': '{}'.format(pie_label_col)})
        
        plt.legend(loc='center', bbox_to_anchor=(1.2, 0.5), labels=final['name'], frameon=False)
        
         
        p = plt.gcf()
        
        plt.setp(pie, width=0.5)
        # Getting current time for filenames
        current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
        
        # Save charts
        if sumavg == 'Average':
            p.savefig('{0}pie_chart_avg_{1}_{2}.png'.format(filepath, col, current_time), transparent=True)
        else:
            p.savefig('{0}pie_chart_{1}_{2}_{3}.png'.format(filepath, product, col, current_time), transparent=True)


def pretty_table(df, b, g, p, q, question_text):
    df = go.Figure(data=go.Table(
        header=dict(height=50,
                    values=['NAME', 'BEFORE', 'AFTER', 'GROWTH'],
                    font = dict(size=20, color='#2d2e83', family='Gotham Black'),
                    fill_color='white',
                    align='center'),
        cells=dict(height=30,
                   values=[df['name'], df['before'], df['after'],
                           df['growth']],
                   font = dict(size=14, family='Gotham'),
                   fill=dict(color=['#f6f6fa',
                                    '#f6f6fa',
                                    '#f6f6fa',
                                    ['#F33447' if '-' in val else '#f6f6fa' for val in df['growth']]]),
                   align='center')))
    
    p = p.replace("'", "") if p != 0 else p
    
    if b == True and g == True:
        gender = 'Boys/Girls'
    elif b == True:
        gender = 'Boys'
    elif g == True:
        gender = 'Girls'
    
    if p == 0:
        df.update_layout(title=dict(text='Ages {0}-{1}, {2}'.format(age1, age2, gender),
                                   font=dict(size=30, color='#2d2e83', family='Gotham Black')),
                        margin=dict(l=0, r=0, b=0, t=50))
    else:
        df.update_layout(title=dict(text='{0} - Ages {1}-{2}, {3}'.format(p, age1, age2, gender),
                                       font=dict(size=30, color='#2d2e83', family='Gotham Black')),
                            margin=dict(l=0, r=0, b=0, t=50))
    
    return [st.header('{} - {}'.format(q, question_text)),
            st.write(''),
            st.write(''),
            st.write(df),
            st.write(''),
            st.write(''),
            st.write('')]




def get_data(olddate1, olddate2, newdate1, newdate2, age1, age2, b, g,
    product, sumavg, qcode, subq, offline, digital, selnames, top,
    bottom, sign, threshold, sort, asc, graph, graph_orient, text_colour,
    month_range, pie_label_col, by_quarter, smooth, ticksize,
    graph_title, title_font, title_fontsize, xaxis, yaxis, label_font,
    label_fontsize, tick_rotation, halo, answer, comparison,
    sub_comparison):
    

    if graph_title == None:
        graph_title = ' '
    if xaxis == None:
        xaxis = ' '
    if yaxis == None:
        yaxis = ' '

        
    # List of all regions
    regions = ['Kids Australia', 'Kids Brazil', 'Kids Canada', 'Kids China',
                'Kids France', 'Kids Germany', 'Kids India', 'Kids Indonesia',
                'Kids Italy', 'Kids Japan', 'Kids South Korea', 'Kids Mexico',
                'Kids Philippines', 'Kids Poland', 'Kids Russia', 'Kids Spain',
                'Kids UK', 'Kids USA']
    
    # text_query = """select top(1) [question_code]
    # ,[question_text]
    # from [dat].[t_question_codes]
    # where question_code = '{}' """.format(q)
    
    # Query_text = pd.read_sql(text_query, cnxn)
    # text = pd.DataFrame(Query_text)
    # question_text = text['question_text'].values[0]
        
        
    # List of all question codes
    qcodes = ['AGE002', 'CHAR001', 'CHAR002', 'CHAR003', 'DBRND001', 'DVCE001', 'DVCE002',
              'DVCE003', 'DVCE004', 'DVCE006', 'DVCE008', 'DVCE009', 'DVCE010',
              'DVCE011', 'DVCE014', 'DVCE015', 'DVCE016', 'DVCE017', 'DVCE018',
              'DVCE019', 'DVCE020', 'DVCE021', 'FILM001', 'FILM002', 'FILM003',
              'FILM004', 'FILM006', 'FILM008', 'FILM009', 'FVRT001', 'FVRT002',
              'FVRT003', 'FVRT004', 'FVRT005', 'FVRT006A', 'FVRT006C', 'FVRT008',
              'FVRT010', 'FVRT012', 'FVRT013', 'FVRT014', 'FVRT015', 'FVRT016',
              'FVRT019', 'FVRT020', 'FVRT021', 'FVRT023', 'FVRT025', 'FVRT050',
              'MUSC001', 'MUSC002', 'MUSC003', 'MUSC004', 'MUSC005', 'PRIV001',
              'PROF000', 'PROF001', 'PROF001A', 'PROF001B', 'PROF001C',
              'PROF001D', 'PROF002', 'PROF003', 'PROF004', 'PROF005', 'FVRT009',
              'PROF006', 'PROF007', 'PROF009', 'PROF013', 'PROF100', 'SAFE005',
              'TEXT002', 'TVOD001', 'TVOD002', 'TVOD003', 'TVOD004', 'TVOD005',
              'TVOD006', 'TVOD007', 'TVOD009', 'TVOD010', 'TVOD012', 'TVOD013',
              'TVOD014', 'TVOD018', 'TVOD019', 'TVOD020', 'YTUB001', 'YTUB002',
              'YTUB003', 'YTUB004', 'YTUB005', 'YTUB006', 'YTUB008', 'YTUB009',
              'AWAR001', 'AWAR002', 'BRND001', 'BRND002', 'BRND004', 'BRND008',
              'BRND009', 'BRND010', 'COOL012', 'COOL100', 'FOOD001', 'FOOD002',
              'FOOD003', 'FOOD004', 'FOOD006', 'FOOD007', 'FOOD010', 'FOOD021',
              'FOOD022', 'FOOD023', 'FOOD024', 'FOOD025', 'FOOD026', 'FOOD028',
              'GAME001', 'GAME003', 'GAME010', 'GAME012', 'GAME013', 'HOBB001',
              'HOBB002', 'HOBB004', 'HOBB008', 'HOBB009', 'HOBB011', 'HOBB012',
              'HOBB013', 'MONY003', 'MONY005', 'MONY006', 'MONY010',
              'MONY011', 'MONY012', 'MONY013', 'ODEV001', 'ODEV002', 'PROF011',
              'PROF012', 'READ001', 'READ002', 'READ003', 'READ004', 'READ006',
              'READ007', 'READ008', 'READ009', 'READ010', 'SHOP001', 'SHOP002',
              'SHOP004', 'SHOP009', 'SHOP010', 'SHOP011', 'SHOP012', 'SHOP015',
              'SHOP016', 'SHOP017', 'SHOP021', 'SHOP022', 'SHOP024', 'TIME001',
              'TIME002', 'TIME004', 'TIME005', 'TIME006', 'FOOD027', 'PROF014',
              'TVOD021', 'HOBB007', 'AGEX001', 'DVCE005', 'DVCE007', 'FVRT011', 
              'TVOD011', 'READ005', 'SHOP023', 'AGE001', 'SAQU002', 'SAQU003',
              'TEXT003', 'TVOD022', 'MOSAICHI', 'MOSAICMG', 'MOSAICMT',
              'MOSAICNC']
    
    
    
    
    # Spellchecker for products
    products = []
    # split up products
    product = product.split(',')
    for p in product:
        p = process.extractOne(p, regions)[0]
        products.append(p)
    
    # For p in products, surround p by quotes ready for the SQL query
    if sumavg == 'Average':
        for p in range(len(products)):
            products[p] = '{}'.format(products[p])
    else:
        for p in range(len(products)):
            products[p] = "'{}'".format(products[p])
    
    
    # If length of products is 1, take product out of list, else convert to
    # tuple
    if len(products) == 1:
        product_string = products[0]   
    
    products_tup = tuple(products)

    
    if b == True and g == True:
        gender = "('Boy', 'Girl')"
    elif b == True:
        gender = "('Boy')"
    elif g == True:
        gender = "('Girl')"
    
    # st.write(sumavg)
    
    # Changing asc variable to True and False    
    if asc == 'Yes':
        asc = True
    else:
        asc = False     
     
    
    if offline:
        survey = "('offline')"
    elif digital:
        survey = "('digital')"
    elif offline and digital:
        survey = "('offline', 'digital')"
    
    
    # Lists to store queries for later use
    old_nums = []
    new_nums = []
    s_nums = []
    
    
    
    # Spellchecking question codes
    qcode = qcode.split(',')
    qs = []
    for q in qcode:
        qs.append(process.extractOne(q, qcodes)[0])
    
    
    if subq != '':
        subqs = subq.split(',')
    else:
        subqs = subq
    
    try:
        st.write('try')
        for q in qs:
            st.write(q)
            for subq in subqs:
                # If question code is FVRT006A OR FVRT006C, change to FVRT006
                # if q in ['FVRT006A', 'FVRT006C']:
                #     q = 'FVRT006'
                
                # If user has selected average
                if sumavg == 'Average':
                    st.write('Going through first average loop')
                    # Writing query for the numerator
                    query_num = """select prof.[record_id] as profile_record_id
                    ,ans.[record_id] as answer_record_id
                    ,[product]
                    ,[date_submitted]
                    ,[age]
                    ,[gender]
                    ,[location]
                    ,[question_code]
                    ,[answer]
                    ,[subquestion]
                    ,[survey_type]
                    from dat.t_profile as prof
                    LEFT JOIN  dat.t_answers  as ans ON (prof.record_id = ans.record_id and question_code='{0}')
                    WHERE prof.date_submitted between '2017/07/01' and '2030/10/31'
                    and gender in {1}
                    and age between {2} and {3}
                    and survey_type in {4}
                    and product in {5} """.format(q if q not in ['FVRT006A', 'FVRT006C'] else 'FVRT006', gender, age1,
                                                  age2, survey,
                                                  "('{}')".format(product_string) if len(products) == 1 else products_tup)
                    
                    # Writing query for the denominator
                    query_denom = """SELECT [record_id]
                    , [date_submitted]
                    , [product]
                    ,[survey_type]
                    ,[gender]
                    FROM [dat].[t_profile]
                    WHERE gender in {0}
                    and age between {1} and {2}
                    and date_submitted between '2017/07/01' and '2030/10/31'
                    and survey_type in {3}
                    and product in {4} """.format(gender, age1, age2, survey,
                                                  "('{}')".format(product_string) if len(products) == 1 else products_tup)
                    
                    # Getting correct question text for inputted question code
                    text_query = """select top(1) [question_code]
                    ,[question_text]
                    from [dat].[t_question_codes]
                    where question_code = '{}' """.format(q if q not in ['FVRT006A', 'FVRT006C'] else 'FVRT006')
                    
                    Query_text = pd.read_sql(text_query, cnxn)
                    text = pd.DataFrame(Query_text)
                    question_text = text['question_text'].values[0]
                        
                        
                    # Putting numerator into a dataframe
                    Query_num = pd.read_sql(query_num, cnxn)
                    num = pd.DataFrame(Query_num)
                    num['date'] = pd.to_datetime(num['date_submitted'])
                    num = num.set_index(num['date'])
                    num = num.sort_index()
                    num = num.replace(['None','2BCLEANED'], np.nan)
                    
                    
                    # Putting denominator into a dataframe
                    Query_denom = pd.read_sql(query_denom, cnxn)
                    denom = pd.DataFrame(Query_denom)
                    denom['date'] = pd.to_datetime(denom['date_submitted'])
                    denom = denom.set_index(denom['date'])
                    denom = denom.sort_index()
                    denom = denom.replace(['None','2BCLEANED'], np.nan)
                    
                    
                    # Selecting dates within the dataframe
                    old_num = num[olddate1:olddate2]
                    old_denom = denom[olddate1:olddate2]
                    new_num = num[newdate1:newdate2]
                    new_denom = denom[newdate1:newdate2]
                    
                    
                    # If there is a subquestion, pass it through the subquestions function
                    if subq != '':
                        old_num, new_num = subquestions(old_num, new_num, subq)
                    
                    # If gender is Boy/Girl, set up conditions for the hued gender graph
                    # if gender == "('Boy', 'Girl')":
                    # # Need to manipulate dataframes based on question code the user has
                    # # selected
                    #     if q == 'MONY011':
                    #         boy = money011(old_num.loc[old_num['gender'] == 'Boy'], new_num.loc[new_num['gender'] == 'Boy'])
                    #         boy['gender'] = 'Boy'
                    #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                            
                            
                    #         girl = money011(old_num.loc[old_num['gender'] == 'Girl'], new_num.loc[new_num['gender'] == 'Girl'])
                    #         girl['gender'] = 'Girl'
                    #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                            
                    #         boygirl = pd.concat([boy, girl], ignore_index=True)
                    #         boygirl['before'] = boygirl['before'].apply(lambda x: round(x,2))
                    #         boygirl['after'] = boygirl['after'].apply(lambda x: round(x,2))
                        
                    #     elif q == 'TIME005':
                    #         boy = time005(old_num.loc[old_num['gender'] == 'Boy'], new_num.loc[new_num['gender'] == 'Boy'])
                    #         boy['gender'] = 'Boy'
                    #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                            
                            
                    #         girl = time005(old_num.loc[old_num['gender'] == 'Girl'], new_num.loc[new_num['gender'] == 'Girl'])
                    #         girl['gender'] = 'Girl'
                    #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                            
                    #         boygirl = pd.concat([boy, girl], ignore_index=True)
                        
                    #     else:
                    #         if q in ['FVRT006A', 'FVRT006C']:
                    #                 old_num, new_num = fvrt006(old_num, new_num, q)
                            
                    #         # Count boy answers
                    #         boy_before = old_num['answer'].loc[old_num['gender'] == 'Boy'].value_counts()
                    #         boy_before = boy_before.apply(lambda x: x/len(old_denom.loc[old_denom['gender'] == 'Boy']))
                    #         boy_after = new_num['answer'].loc[new_num['gender'] == 'Boy'].value_counts()
                    #         boy_after = boy_after.apply(lambda x: x/len(new_denom.loc[new_denom['gender'] == 'Boy']))
                            
                    #         # Convert the dataframes
                    #         boy_before = pd.DataFrame({'name':boy_before.index, 'value':boy_before.values})
                    #         boy_after = pd.DataFrame({'name':boy_after.index, 'value':boy_after.values})
                            
                    #         # Merge before and after dataframes together
                    #         boy = pd.merge(boy_before, boy_after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                            
                    #         # Selecting where before is above zero to prevent dividebyzero error
                    #         boy = boy.loc[boy['before'] > 0]
                            
                    #         # Calculating growth
                    #         boy['growth'] = boy.apply(lambda x: (x['after'] / 
                    #          x['before'])-1, axis=1)
                            
                    #         # Multiply by 100 to get growth percentages           
                    #         boy['growth'] = boy['growth'].apply(lambda x: x*100)
                            
                            
                    #         # Making a gender column so we can hue on it later
                    #         boy['gender'] = 'Boy'
                            
                    #         # Using the table_type function above
                    #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                            
                            
                            
                    #         # Count girl answers
                    #         girl_before = old_num['answer'].loc[old_num['gender'] == 'Girl'].value_counts()
                    #         girl_before = girl_before.apply(lambda x: x/len(old_denom.loc[old_denom['gender'] == 'Girl']))
                    #         girl_after = new_num['answer'].loc[new_num['gender'] == 'Girl'].value_counts()
                    #         girl_after = girl_after.apply(lambda x: x/len(new_denom.loc[new_denom['gender'] == 'Girl']))
                            
                    #         # Convert to dataframes
                    #         girl_before = pd.DataFrame({'name':girl_before.index, 'value':girl_before.values})
                    #         girl_after = pd.DataFrame({'name':girl_after.index, 'value':girl_after.values})
                            
                    #         # Merge before and after dataframes together
                    #         girl = pd.merge(girl_before, girl_after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                            
                    #         # Selecting where before is above zero to prevent dividebyzero error
                    #         girl = girl.loc[girl['before'] > 0]
                            
                    #         # Calculating growth
                    #         girl['growth'] = girl.apply(lambda x: (x['after'] / 
                    #         x['before'])-1, axis=1)
                            
                    #         # Multiply by 100 to get growth percentages           
                    #         girl['growth'] = girl['growth'].apply(lambda x: x*100)
                            
                            
                    #         # Making a gender column so we can hue on it later
                    #         girl['gender'] = 'Girl'
                            
                    #         # Using the table_type function above
                    #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                            
                    #         # Concatenating the boy and girl tables into one dataframe
                    #         boygirl = pd.concat([boy, girl], ignore_index=True)
                            
                    #         # Multiply before and after by 100 and round to 2 decimal places
                    #         boygirl['before'] = boygirl['before'].apply(lambda x: round(x*100,2))
                    #         boygirl['after'] = boygirl['after'].apply(lambda x: round(x*100,2))
                    
                    # Appending old num and new num to the lists above
                    old_nums.append(old_num)
                    new_nums.append(new_num)
                    #print(num['answer'])
                    
                    # Range to collect sample from
                    s_num = num[olddate1:newdate2]
                    
                    # Sample size is the length of the s_num dataframe
                    sample = len(s_num)
                    
                    # Append s_num to the s_nums list
                    s_nums.append(s_num)
                    
                    # Getting total surveyed from old data
                    sample_old = len(old_denom)
                    
                    
                    # If total surveyed is zero, print invalid date range
                    if sample_old == 0:
                        print('failed')
                        # sheet['{}'.format(loc)].offset(-1,0).value = ''
                        # sheet["{}".format(loc)].value = 'No data for this region!'
                    else:
                        # If question code is one of the dodgy question codes, use the 
                        # functions for them above
                        if q == 'MONY011':
                            df_merge_col = money011(old_num, new_num)
                        elif q == 'TIME005':
                            df_merge_col = time005(old_num, new_num)
                        else:
                            if q in ['FVRT006A', 'FVRT006C']:
                                if b == False and g == False:
                                    old_num, new_num = fvrt006(old_num, new_num, q)
                            
                            # Do value counts for old data
                            r_old = old_num['answer'].value_counts()
                            
                            # Getting percentage values that match the portal
                            r_old = r_old.apply(lambda x: x/sample_old)
                            
                            # Getting sample of new data
                            sample_new = len(new_denom)
                            
                            # getting value counts of new data
                            r_new = new_num['answer'].value_counts()
                            
                            # Getting percentage values that match the portal
                            r_new = r_new.apply(lambda x: x/sample_new)
                            
                            # Making dataframe to display values for each time period
                            before = pd.DataFrame({'name':r_old.index, 'value':r_old.values})
                            after = pd.DataFrame({'name':r_new.index, 'value':r_new.values})
                            
                            # Mergeing the dataframes ready for comparison
                            df_merge_col = pd.merge(before, after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                            df_merge_col = df_merge_col.loc[df_merge_col['before'] > 0]
                            
                            # Calculating growth
                            df_merge_col['growth'] = df_merge_col.apply(lambda x: (x['after'] / 
                                                                                    x['before'])-1, axis=1)
                            
                            # Scaling by 100 and sorting from largest to smallest            
                            df_merge_col['growth'] = df_merge_col['growth'].apply(lambda x: x*100)
                        
                        # Sort values by filters the user has inputted
                        df_merge_col = df_merge_col.sort_values(by=sort, ascending=asc)
                        
                        # If question code is TIME005, get a copy of the dataframe
                        if q == 'TIME005':
                            # Displaying time in another format, making use of the hour_mins
                            # function above
                            time = df_merge_col.copy()
                            time['before'] = time['before'].apply(lambda x: "'{}'".format(str(datetime.timedelta(hours=x))))
                            time['before'] = time['before'].apply(lambda x: hour_mins(x))
                            time['after'] = time['after'].apply(lambda x: "'{}'".format(str(datetime.timedelta(hours=x))))
                            time['after'] = time['after'].apply(lambda x: hour_mins(x))
                            
                        # Creating dataframe in the style the user wants
                        final = table_type(df_merge_col, selnames, top, bottom, threshold, sign, sort)
                        
                        # If q is MONY011 round before and after columns to 2 decimal places 
                        if q == 'MONY011':
                            final['before'] = final['before'].apply(lambda x: round(x,2))
                            final['after'] = final['after'].apply(lambda x: round(x,2))
                        
                        # If code is TIME005 do nothing
                        elif q == 'TIME005':
                            final = final
                        else:
                            # Else multiply before and after columns by 100 and round to 
                            # 2 decimal places
                            final['before'] = final['before'].apply(lambda x: round(x*100,2))
                            final['after'] = final['after'].apply(lambda x: round(x*100,2))
                        
                        # Round growth column to nearest integer
                        final['growth'] = final['growth'].apply(lambda x: int(round(x)))
                        
                        # If products_tup isn't a tuple, replace the brackets with empty strings
                        if isinstance(products_tup, tuple) == False:
                            p_final = products_tup.replace("('", "")
                            p_final = p_final.replace("')", "")
                        else:
                            # Else make p_final equal to products_tup
                            p_final = products_tup
                        
                        
                        # Colour wheel
                        colors = ['#96D5EE', '#27AAE1', '#2D2E83', '#7277C7',
                          '#00A19A', '#ACCC00', '#662483', '#9D96C9',
                          '#CE2F6C', '#FAA932', '#464B9A', '#347575',
                          '#6AC2C2', '#76AF41', '#A33080', '#EB8EDB',
                          '#F66D9B', '#D95D27', '#F7915D', '#FFC814']  
                        
                        # # If display is True
                        # if display == True:
                        # If user has set graphs to Yes
                        if graph == 'Bar Chart':
                            # count starting at 1
                            c = 1
                            for col in ['before', 'after', 'growth']:
                                # if gender == "('Boy', 'Girl')":
                                #     try:
                                #         # Setting user inputted text colour
                                #         plt.rcParams['text.color'] = text_colour
                    
                                #         # Vertical hue graph
                                #         if graph_orient == 'vertical':
                                #             graph = sns.catplot(x='name', y=col,
                                #             data=boygirl,
                                #             kind='bar', palette=['#D95D27', '#662483'],
                                #             hue='gender',
                                #             height=6, aspect=2, legend=False,
                                #             orient='v')
                                #         else:
                                #             # Horizontal hue graph
                                #             graph = sns.catplot(x=col, y='name',
                                #             data=boygirl,
                                #             kind='bar', palette=['#D95D27', '#662483'],
                                #             hue='gender',
                                #             height=6, aspect=2, legend=False,
                                #             orient='h')
                    
                    
                                #         # Fonts for title and labels
                                #         tfont = {'fontname':'{}'.format(title_font)}
                                #         lfont = {'fontname':'{}'.format(label_font)}
                    
                                #         # Set user inputted title with correct formating
                                #         plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
                    
                                #         # Adding correct formating to graph labels
                                #         if graph_orient == 'vertical':
                                #             plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                #             plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                #         else:
                                #             # Swapping axis labels for horizontal graph
                                #             plt.xlabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                #             plt.ylabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour) 
                    
                                #         # Rotate x-axis ticks based on what user has
                                #         # selected for tick_rotation
                                #         plt.xticks(rotation=tick_rotation, fontsize=ticksize)
                                #         plt.yticks(fontsize=ticksize)
                                        
                                #         ax = plt.gca()
                                        
                                #         if graph_orient == 'vertical':
                                #             ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                #         else:
                                #             ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                        
                                #         # If halo is set to On
                                #         if halo == 'On':
                                #             if q == 'TIME005':
                                #                 if c != 3:
                                #                     # If count = 3 don't covert to hours and mins
                                #                     bar_labels(ax, time='y', orientation=graph_orient)
                                #                 else:
                                #                     bar_labels(ax, orientation=graph_orient)
                                #             else:
                                #                 bar_labels(ax, orientation=graph_orient)
                                        
                                #         # Setting colour of the tick labels
                                #         [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
                                #         [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
                                #         [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
                                #         [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
                                        
                                #         # Adding legend
                                #         plt.legend(bbox_to_anchor=(0.6, -0.1), ncol=2, frameon=False).set_title('')
                                        
                                        
                                #         # Despine graph
                                #         sns.despine(left=True)
                                        
                                #         # Getting current time for filenames
                                #         current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
                                    
                                #     except:
                                #         print('failed')
                                #         # If there was an error with the hue graphs,
                                #         # print this message
                                #         # sheet["A38"].value = 'Could not create hue graph'
                    
                                # Main graphs
                                try:
                                    # Setting user inputted text colour
                                    plt.rcParams['text.color'] = text_colour
                                    
                                    # Vertical graph
                                    if graph_orient == 'vertical':
                                        graph = sns.catplot(x='name', y=col,
                                        data=final,
                                        kind='bar', palette=colors,
                                        height=6, aspect=2, legend=False,
                                        orient='v')
                                    
                                    # Horizontal graph
                                    else:
                                        graph = sns.catplot(x=col, y='name',
                                        data=final,
                                        kind='bar', palette=colors,
                                        height=6, aspect=2, legend=False,
                                        orient='h')
                                    
                                    # Fonts for title and labels
                                    tfont = {'fontname':'{}'.format(title_font)}
                                    lfont = {'fontname':'{}'.format(label_font)}
                                    
                                    # Set user inputted title with correct formating
                                    plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
                                    
                                    
                                    # Adding correct formating to graph labels
                                    if graph_orient == 'vertical':
                                        plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                        plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                    else:
                                        # Swapping axis labels for horizontal graph
                                        plt.xlabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                        plt.ylabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour) 
                                    
                                    # Rotate x-axis ticks based on what user has
                                    # selected for tick_rotation
                                    plt.xticks(rotation=tick_rotation, fontsize=ticksize)
                                    plt.yticks(fontsize=ticksize)
                                    
                                    ax = plt.gca()
                                    
                                    if graph_orient == 'vertical':
                                        ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                    else:
                                        ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                    
                                    # If halo is set to On
                                    if halo == 'On':
                                        if q == 'TIME005':
                                            if c != 3:
                                                # If count = 3 don't covert to hours and mins
                                                bar_labels(ax, time='y', orientation=graph_orient)
                                            else:
                                                bar_labels(ax, orientation=graph_orient)
                                        else:
                                            bar_labels(ax, orientation=graph_orient)
                                    
                                    
                                    # Setting colour of the tick labels       
                                    [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
                                    [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
                                    [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
                                    [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
                                    
                                    # Despine graphs
                                    sns.despine(left=True)
                                    
                                    # Getting current time for filenames
                                    current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
                                    
                                    
                                    # Add 1 to the count
                                    c += 1
                                except:
                                    print('failed')
                                    # sheet["A38"].value = 'Invalid colour'
                    
                        elif graph == 'Line Graph':
                            try:
                                line_graph(by_quarter, smooth, newdate2, num, denom, month_range, product, sumavg, selnames, text_colour, graph_title, title_font, title_fontsize, xaxis, yaxis, label_font, label_fontsize, tick_rotation, ticksize)
                            except:
                                print('failed')
                                # sheet["A38"].value = 'Could not create graph'
                        elif graph == 'Pie Chart':
                            try:
                                pie_chart(final, sumavg, product, text_colour, graph_title, title_fontsize, title_font, pie_label_col)
                            except:
                                print('failed')
                                # sheet["A38"].value = 'Could not create graph'
                    
                    
                        # If question code isn't in MONY011 or TIME005 then add % sign 
                        # to before and after columns
                        if q not in['MONY011', 'TIME005']:
                            final['before'] = final['before'].apply(lambda x: str(x) + '%')
                            final['after'] = final['after'].apply(lambda x: str(x) + '%')
                        
                        # Add % sign to the growth column
                        final['growth'] = final['growth'].apply(lambda x: str(x) + '%')
                        
                        
                        # If question code is TIME005 set table style and round the growth
                        # column to the nearest integer, add a % sign
                        if q == 'TIME005':
                            final = table_type(time, selnames, top, bottom, threshold, sign, sort)
                            final['growth'] = final['growth'].apply(lambda x: str(int(round(x))) + '%')
                         
                    # p = 0
                    # pretty_table(final, b, g, p, q, question_text)
                        
                # If user has selected Sum             
                else:
                    st.write('Going through first sum loop')
                    # For every product the user has inputted    
                    for p in products_tup:
                        # Writing query for the numerator
                        query_num = """select prof.[record_id] as profile_record_id
                        ,ans.[record_id] as answer_record_id
                        ,[product]
                        ,[date_submitted]
                        ,[age]
                        ,[gender]
                        ,[location]
                        ,[question_code]
                        ,[answer]
                        ,[subquestion]
                        ,[survey_type]
                        from dat.t_profile as prof
                        LEFT JOIN  dat.t_answers  as ans ON (prof.record_id = ans.record_id and question_code='{0}')
                        WHERE prof.date_submitted between '2017/07/01' and '2030/10/31'
                        and gender in {1}
                        and age between {2} and {3}
                        and survey_type in {4}
                        and product = {5} """.format(q if q not in ['FVRT006A', 'FVRT006C'] else 'FVRT006', gender, age1,
                                                      age2, survey,
                                                      product_string if len(products_tup) == 1 else p)
                        
        
                        # Writing query for the denominator
                        query_denom = """SELECT [record_id]
                        , [date_submitted]
                        , [product]
                        ,[survey_type]
                        ,[gender]
                        FROM [dat].[t_profile]
                        WHERE gender in {0}
                        and age between {1} and {2}
                        and date_submitted between '2017/07/01' and '2030/10/31'
                        and survey_type in {3}
                        and product = {4} """.format(gender, age1, age2, survey,
                                                      product_string if len(products_tup) == 1 else p)
    
                        # Getting correct question text for inputted question code
                        text_query = """select top(1) [question_code]
                        ,[question_text]
                        from [dat].[t_question_codes]
                        where question_code = '{}' """.format(q if q not in ['FVRT006A', 'FVRT006C'] else 'FVRT006')
                        
                        Query_text = pd.read_sql(text_query, cnxn)
                        text = pd.DataFrame(Query_text)
                        question_text = text['question_text'].values[0]
                          
                        # Putting numerator into a dataframe
                        Query_num = pd.read_sql(query_num, cnxn)
                        num = pd.DataFrame(Query_num)
                        num['date'] = pd.to_datetime(num['date_submitted'])
                        num = num.set_index(num['date'])
                        num = num.sort_index()
                        num = num.replace(['None','2BCLEANED'], np.nan)
                         
                        
                        # Putting denominator into a dataframe
                        Query_denom = pd.read_sql(query_denom, cnxn)
                        denom = pd.DataFrame(Query_denom)
                        denom['date'] = pd.to_datetime(denom['date_submitted'])
                        denom = denom.set_index(denom['date'])
                        denom = denom.sort_index()
                        denom = denom.replace(['None','2BCLEANED'], np.nan)
                        
                        
                        # Selecting dates within the dataframe
                        old_num = num[olddate1:olddate2]
                        old_denom = denom[olddate1:olddate2]
                        new_num = num[newdate1:newdate2]
                        new_denom = denom[newdate1:newdate2]
                        
                        # If there is a subquestion, pass it through the subquestions function
                        if subq != None:
                            old_num, new_num = subquestions(old_num, new_num, subq)
                        
                        # If gender is Boy/Girl, set up conditions for the hued gender graph
                        # if gender == "('Boy', 'Girl')":
                        #     # Need to manipulate dataframes based on question code the user has
                        #     # selected
                        #     if q == 'MONY011':
                        #         boy = money011(old_num.loc[old_num['gender'] == 'Boy'], new_num.loc[new_num['gender'] == 'Boy'])
                        #         # boy = boy.sort_values(by=sort, ascending=asc)
                        #         boy['gender'] = 'Boy'
                        #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                                
                                
                        #         girl = money011(old_num.loc[old_num['gender'] == 'Girl'], new_num.loc[new_num['gender'] == 'Girl'])
                        #         # girl = girl.sort_values(by=sort, ascending=asc)
                        #         girl['gender'] = 'Girl'
                        #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                                
                        #         boygirl = pd.concat([boy, girl], ignore_index=True)
                        #         boygirl['before'] = boygirl['before'].apply(lambda x: round(x,2))
                        #         boygirl['after'] = boygirl['after'].apply(lambda x: round(x,2))
                            
                        #     elif q == 'TIME005':
                        #         boy = time005(old_num.loc[old_num['gender'] == 'Boy'], new_num.loc[new_num['gender'] == 'Boy'])
                        #         # boy = boy.sort_values(by=sort, ascending=asc)
                        #         boy['gender'] = 'Boy'
                        #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                                
                                
                        #         girl = time005(old_num.loc[old_num['gender'] == 'Girl'], new_num.loc[new_num['gender'] == 'Girl'])
                        #         # girl = girl.sort_values(by=sort, ascending=asc)
                        #         girl['gender'] = 'Girl'
                        #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                                
                        #         boygirl = pd.concat([boy, girl], ignore_index=True)
                            
                        #     else:
                        #         if q in ['FVRT006A', 'FVRT006C']:
                        #             old_num, new_num = fvrt006(old_num, new_num, q)
                                
                                
                        #         # Count boy answers    
                        #         boy_before = old_num['answer'].loc[old_num['gender'] == 'Boy'].value_counts()
                        #         boy_before = boy_before.apply(lambda x: x/len(old_denom.loc[old_denom['gender'] == 'Boy']))
                        #         boy_after = new_num['answer'].loc[new_num['gender'] == 'Boy'].value_counts()
                        #         boy_after = boy_after.apply(lambda x: x/len(new_denom.loc[new_denom['gender'] == 'Boy']))
                                
                        #         # Convert to dataframes
                        #         boy_before = pd.DataFrame({'name':boy_before.index, 'value':boy_before.values})
                        #         boy_after = pd.DataFrame({'name':boy_after.index, 'value':boy_after.values})
                                
                        #         # Merge dataframes
                        #         boy = pd.merge(boy_before, boy_after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                        #         boy = boy.loc[boy['before'] > 0]
                                
                        #         # Calculating growth
                        #         boy['growth'] = boy.apply(lambda x: (x['after'] / 
                        #          x['before'])-1, axis=1)
                                
                        #         # Multiplying growth column by 100           
                        #         boy['growth'] = boy['growth'].apply(lambda x: x*100)
                                
                        #         # # Sort values based on filters the user has selected
                        #         # boy = boy.sort_values(by=sort, ascending=asc)
                                
                        #         # Setting a gender column equal to Boy
                        #         boy['gender'] = 'Boy'
                                
                        #         # Using the table_type function above
                        #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                                
                                
                                
                        #         # Count girl answers
                        #         girl_before = old_num['answer'].loc[old_num['gender'] == 'Girl'].value_counts()
                        #         girl_before = girl_before.apply(lambda x: x/len(old_denom.loc[old_denom['gender'] == 'Girl']))
                        #         girl_after = new_num['answer'].loc[new_num['gender'] == 'Girl'].value_counts()
                        #         girl_after = girl_after.apply(lambda x: x/len(new_denom.loc[new_denom['gender'] == 'Girl']))
                                
                        #         # Convert the dataframes
                        #         girl_before = pd.DataFrame({'name':girl_before.index, 'value':girl_before.values})
                        #         girl_after = pd.DataFrame({'name':girl_after.index, 'value':girl_after.values})
                                
                        #         # Merge dataframes
                        #         girl = pd.merge(girl_before, girl_after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                        #         girl = girl.loc[girl['before'] > 0]
                                
                        #         # Calculating growth
                        #         girl['growth'] = girl.apply(lambda x: (x['after'] / 
                        #         x['before'])-1, axis=1)
                                
                        #         # Multiplying growth column by 100            
                        #         girl['growth'] = girl['growth'].apply(lambda x: x*100)
                                 
                        #         # # Sort values based on filters the user has selected 
                        #         # girl = girl.sort_values(by=sort, ascending=asc)
                                
                        #         # Setting a gender column equal to Girl
                        #         girl['gender'] = 'Girl'
                                
                        #         # Using the table_type function above
                        #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                                
                        #         # Concatenating the boy and girl tables into one dataframe
                        #         boygirl = pd.concat([boy, girl], ignore_index=True)
                                
                        #         # Multiply before and after by 100 and round to 2 decimal places
                        #         boygirl['before'] = boygirl['before'].apply(lambda x: round(x*100,2))
                        #         boygirl['after'] = boygirl['after'].apply(lambda x: round(x*100,2))
                        
                        
                        # Appending old num and new num to the lists above
                        old_nums.append(old_num)
                        new_nums.append(new_num)
                        
                        # Range to collect sample from
                        s_num = num[olddate1:newdate2]
                        
                        # Sample size is the length of the s_num dataframe
                        sample = len(s_num)
                        
                        # Append s_num to the s_nums list
                        s_nums.append(s_num)
                        
                        # Getting total surveyed from old data
                        sample_old = len(old_denom)
                        
                        
                        # If total surveyed is zero, print invalid date range
                        if sample_old == 0:
                            print('failed')
                            # sheet['{}'.format(loc)].offset(-1,0).value = ''
                            # sheet["{}".format(loc)].value = 'No data for this region!'
                        else:
                            # If question code is one of the dodgy question codes, use the 
                            # functions for them above
                            if q == 'MONY011':
                                df_merge_col = money011(old_num, new_num)
                            elif q == 'TIME005':
                                df_merge_col = time005(old_num, new_num)
                            else:
                                if q in ['FVRT006A', 'FVRT006C']:
                                    if b == False and g == False:
                                        old_num, new_num = fvrt006(old_num, new_num, q)
                                        
                                # Do value counts for old data
                                r_old = old_num['answer'].value_counts()
                                
                                # Getting percentage values that match the portal
                                r_old = r_old.apply(lambda x: x/sample_old)
                                
                                # Getting sample of new data
                                sample_new = len(new_denom)
                                
                                # Do value counts for new data
                                r_new = new_num['answer'].value_counts()
             
                                # Getting percentage values that match the portal
                                r_new = r_new.apply(lambda x: x/sample_new)
                                
                                # Making dataframe to display values for each time period
                                before = pd.DataFrame({'name':r_old.index, 'value':r_old.values})
                                after = pd.DataFrame({'name':r_new.index, 'value':r_new.values})
                                # print(old_num)
                                # sleep(100000)
                                # Mergeing the dataframes ready for comparison
                                df_merge_col = pd.merge(before, after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                                df_merge_col = df_merge_col.loc[df_merge_col['before'] > 0]
                                
                                # Calculating growth
                                df_merge_col['growth'] = df_merge_col.apply(lambda x: (x['after'] / 
                                                                                        x['before'])-1, axis=1)
                                
                                # Scaling growth by 100            
                                df_merge_col['growth'] = df_merge_col['growth'].apply(lambda x: x*100)
                            
                            # Sort values by filters the user has inputted
                            df_merge_col = df_merge_col.sort_values(by=sort, ascending=asc)
                            
                            # If question code is TIME005, get a copy of the dataframe
                            if q == 'TIME005':
                                # Displaying time in another format, making use of the hour_mins
                                # function above
                                time = df_merge_col.copy()
                                time['before'] = time['before'].apply(lambda x: "'{}'".format(str(datetime.timedelta(hours=x))))
                                time['before'] = time['before'].apply(lambda x: hour_mins(x))
                                time['after'] = time['after'].apply(lambda x: "'{}'".format(str(datetime.timedelta(hours=x))))
                                time['after'] = time['after'].apply(lambda x: hour_mins(x))
                              
                            # Creating dataframe in the style the user wants
                            final = table_type(df_merge_col, selnames, top, bottom, threshold, sign, sort)
                            
                            # If q is MONY011 round before and after columns to 2 decimal places
                            if q == 'MONY011':
                                final['before'] = final['before'].apply(lambda x: round(x,2))
                                final['after'] = final['after'].apply(lambda x: round(x,2))
                            
                            # If code is TIME005 do nothing
                            elif q == 'TIME005':
                                final = final
                            else:
                                # Else multiply before and after columns by 100 and round to 
                                # 2 decimal places
                                final['before'] = final['before'].apply(lambda x: round(x*100,2))
                                final['after'] = final['after'].apply(lambda x: round(x*100,2))
                            
                            # Round growth column to nearest integer
                            final['growth'] = final['growth'].apply(lambda x: int(round(x)))
                            
                            # Replace apostrophies in p with an empty string
                            p_final = p.replace("'", "")
                            
                            # Colour wheel
                            colors = ['#96D5EE', '#27AAE1', '#2D2E83', '#7277C7',
                                      '#00A19A', '#ACCC00', '#662483', '#9D96C9',
                                      '#CE2F6C', '#FAA932', '#464B9A', '#347575',
                                      '#6AC2C2', '#76AF41', '#A33080', '#EB8EDB',
                                      '#F66D9B', '#D95D27', '#F7915D', '#FFC814']  
                            
                            # # If display is True
                            # if display == True:
                            # If user has set graphs to Yes
                            if graph == 'Bar Chart':
                                # Count starting at 1
                                c = 1
                                for col in ['before', 'after', 'growth']:
                                    # if gender == "('Boy', 'Girl')":
                                    #     try:
                                    #         # Setting user inputted text colour
                                    #         plt.rcParams['text.color'] = text_colour
                                            
                                    #         # Vertical graph
                                    #         if graph_orient == 'vertical':
                                    #             graph = sns.catplot(x='name', y=col,
                                    #             data=boygirl,
                                    #             kind='bar', palette=['#D95D27', '#662483'],
                                    #             hue='gender',
                                    #             height=6, aspect=2, legend=False,
                                    #             orient='v')
                                    #         else:
                                    #             # Horizontal graph
                                    #             graph = sns.catplot(x=col, y='name',
                                    #             data=boygirl,
                                    #             kind='bar', palette=['#D95D27', '#662483'],
                                    #             hue='gender',
                                    #             height=6, aspect=2, legend=False,
                                    #             orient='h')
                                            
                                            
                                    #         # Fonts for title and labels
                                    #         tfont = {'fontname':'{}'.format(title_font)}
                                    #         lfont = {'fontname':'{}'.format(label_font)}
                                            
                                    #         # Set user inputted title with correct formating
                                    #         plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
                                            
                                    #         # Adding correct formatting from graph labels
                                    #         if graph_orient == 'vertical':
                                    #             plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                    #             plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                    #         else:
                                    #             # Swapping axis labels for horizontal graph
                                    #             plt.xlabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                    #             plt.ylabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour) 
                                            
                                    #         # Rotating x-axis ticks based on what the user has selected
                                    #         plt.xticks(rotation=tick_rotation, fontsize=ticksize)
                                    #         plt.yticks(fontsize=ticksize)
                                            
                                    #         ax = plt.gca()
                                            
                                    #         if graph_orient == 'vertical':
                                    #             ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                    #         else:
                                    #             ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                            
                                    #         if halo == 'On':
                                    #             # If halo is set to On
                                    #             if q == 'TIME005':
                                    #                 if c != 3:
                                    #                     # If count = 3, don't convert to hours and mins
                                    #                     bar_labels(ax, time='y', orientation=graph_orient)
                                    #                 else:
                                    #                     bar_labels(ax, orientation=graph_orient)
                                    #             else:
                                    #                 bar_labels(ax, orientation=graph_orient)
                                            
                                    #         # Setting colour of the tick labels
                                    #         [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
                                    #         [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
                                    #         [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
                                    #         [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
                                            
                                    #         # Plotting a legend
                                    #         plt.legend(bbox_to_anchor=(0.6, -0.1), ncol=2, frameon=False).set_title('')
                                            
                                    #         # Despine graphs
                                    #         sns.despine(left=True)
                                            
                                    #         # Getting current time for filenames
                                    #         current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
                                            
                                    #     except:
                                    #         print('failed')
                                    #         # If there was an error with the hue graphs,
                                    #         # print this message
                                    #         # sheet["A38"].value = 'Could not create hue graph'
                        
                                    try:
                                    #Setting user inputted text colour
                                        plt.rcParams['text.color'] = text_colour
                                        
                                        # Vertical graph
                                        if graph_orient == 'vertical':
                                            graph = sns.catplot(x='name', y=col,
                                            data=final,
                                            kind='bar', palette=colors,
                                            height=6, aspect=2, legend=False,
                                            orient='v')
                                        else:
                                            # Horizontal graph
                                            graph = sns.catplot(x=col, y='name',
                                            data=final,
                                            kind='bar', palette=colors,
                                            height=6, aspect=2, legend=False,
                                            orient='h')
                                            
    
                                        # Fonts for title and labels
                                        tfont = {'fontname':'{}'.format(title_font)}
                                        lfont = {'fontname':'{}'.format(label_font)}
                                        
                                        # Set user inputted title with correct formating
                                        plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
                                        
                                        # Adding correct formatting from graph labels
                                        if graph_orient == 'vertical':
                                            plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                            plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                        else:
                                            # Swapping axis labels for horizontal graph
                                            plt.xlabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                            plt.ylabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour) 
                                        
                                        # Rotating x-axis ticks based on what the user has selected
                                        plt.xticks(rotation=tick_rotation, fontsize=ticksize)
                                        plt.yticks(fontsize=ticksize)
                                        
                                        ax = plt.gca()
                                        if graph_orient == 'vertical':
                                            ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                        else:
                                            ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                        
                                        if halo == 'On':
                                            # If halo is set to On
                                            if q == 'TIME005':
                                                if c != 3:
                                                    # If count = 3, don't convert to hours and mins
                                                    bar_labels(ax, time='y', orientation=graph_orient)
                                                else:
                                                    bar_labels(ax, orientation=graph_orient)
                                            else:
                                                bar_labels(ax, orientation=graph_orient)
                                        
                                        # Setting colour of the tick labels        
                                        [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
                                        [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
                                        [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
                                        [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
                                        
                                        # Despine graphs
                                        sns.despine(left=True)
                                        
                                        # Getting current time for filenames
                                        current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
    
                                        # Save graphs
                                        # graph.savefig('{0}{1}_{2}_{3}.png'.format(filepath, p_final, col, current_time), transparent=True)
                                        # Add 1 to count
                                        c += 1
                                    except:
                                        print('failed')
                                        # sheet["A38"].value = 'Invalid colour'
                                    # except Exception as e: print(e)
                                    # sleep(100000)
                        
                            elif graph == 'Line Graph':
                                try:
                                    line_graph(by_quarter, smooth, newdate2, num, denom, month_range, product, sumavg, selnames, text_colour, graph_title, title_font, title_fontsize, xaxis, yaxis, label_font, label_fontsize, tick_rotation, ticksize)
                                except:
                                    print('failed')
                                    # sheet["A38"].value = 'Could not create graph'
                            elif graph == 'Pie Chart':
                                try:
                                    pie_chart(final, sumavg, product, text_colour, graph_title, title_fontsize, title_font, pie_label_col)
                                except:
                                    print('failed')
                                    # sheet["A38"].value = 'Could not create graph'
                            
                            # If question code isn't in MONY011 or TIME005 then add % sign 
                            # to before and after columns
                            if q not in['MONY011', 'TIME005']:
                                final['before'] = final['before'].apply(lambda x: str(x) + '%')
                                final['after'] = final['after'].apply(lambda x: str(x) + '%')
                            
                            # Add % sign to the growth column
                            final['growth'] = final['growth'].apply(lambda x: str(x) + '%')
                            
                            # If question code is TIME005 set table style and round the growth
                            # column to the nearest integer, add a % sign
                            if q == 'TIME005':
                                final = table_type(time, selnames, top, bottom, threshold, sign, sort)
                                final['growth'] = final['growth'].apply(lambda x: str(int(round(x))) + '%')
                           
        pretty_table(final, b, g, p, q, question_text)
        st.write('end of try')
        
    except:
        st.write('exception')
        for q in qs:
        # st.write('help')
        # If question code is FVRT006A OR FVRT006C, change to FVRT006
        # if q in ['FVRT006A', 'FVRT006C']:
        #     q = 'FVRT006'
        
        
            # If user has selected average
            if sumavg == 'Average':
                st.write('Going through second average loop')
                # Writing query for the numerator
                query_num = """select prof.[record_id] as profile_record_id
                ,ans.[record_id] as answer_record_id
                ,[product]
                ,[date_submitted]
                ,[age]
                ,[gender]
                ,[location]
                ,[question_code]
                ,[answer]
                ,[subquestion]
                ,[survey_type]
                from dat.t_profile as prof
                LEFT JOIN  dat.t_answers  as ans ON (prof.record_id = ans.record_id and question_code='{0}')
                WHERE prof.date_submitted between '2017/07/01' and '2030/10/31'
                and gender in {1}
                and age between {2} and {3}
                and survey_type in {4}
                and product in {5} """.format(q if q not in ['FVRT006A', 'FVRT006C'] else 'FVRT006', gender, age1,
                                              age2, survey,
                                              "('{}')".format(product_string) if len(products) == 1 else products_tup)
                
                # Writing query for the denominator
                query_denom = """SELECT [record_id]
                , [date_submitted]
                , [product]
                ,[survey_type]
                ,[gender]
                FROM [dat].[t_profile]
                WHERE gender in {0}
                and age between {1} and {2}
                and date_submitted between '2017/07/01' and '2030/10/31'
                and survey_type in {3}
                and product in {4} """.format(gender, age1, age2, survey,
                                              "('{}')".format(product_string) if len(products) == 1 else products_tup)
                
                # Getting correct question text for inputted question code
                text_query = """select top(1) [question_code]
                ,[question_text]
                from [dat].[t_question_codes]
                where question_code = '{}' """.format(q if q not in ['FVRT006A', 'FVRT006C'] else 'FVRT006')
                
                Query_text = pd.read_sql(text_query, cnxn)
                text = pd.DataFrame(Query_text)
                question_text = text['question_text'].values[0]
                    
                    
                # Putting numerator into a dataframe
                Query_num = pd.read_sql(query_num, cnxn)
                num = pd.DataFrame(Query_num)
                num['date'] = pd.to_datetime(num['date_submitted'])
                num = num.set_index(num['date'])
                num = num.sort_index()
                num = num.replace(['None','2BCLEANED'], np.nan)
                
                
                # Putting denominator into a dataframe
                Query_denom = pd.read_sql(query_denom, cnxn)
                denom = pd.DataFrame(Query_denom)
                denom['date'] = pd.to_datetime(denom['date_submitted'])
                denom = denom.set_index(denom['date'])
                denom = denom.sort_index()
                denom = denom.replace(['None','2BCLEANED'], np.nan)
                
                
                # Selecting dates within the dataframe
                old_num = num[olddate1:olddate2]
                old_denom = denom[olddate1:olddate2]
                new_num = num[newdate1:newdate2]
                new_denom = denom[newdate1:newdate2]
                st.write(query_num)
                
                # If there is a subquestion, pass it through the subquestions function
                if subq != '':
                    old_num, new_num = subquestions(old_num, new_num, subq)
                    print('sgo8yuhaso9ergh')
                # If gender is Boy/Girl, set up conditions for the hued gender graph
                # if gender == "('Boy', 'Girl')":
                # # Need to manipulate dataframes based on question code the user has
                # # selected
                #     if q == 'MONY011':
                #         boy = money011(old_num.loc[old_num['gender'] == 'Boy'], new_num.loc[new_num['gender'] == 'Boy'])
                #         boy['gender'] = 'Boy'
                #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                        
                        
                #         girl = money011(old_num.loc[old_num['gender'] == 'Girl'], new_num.loc[new_num['gender'] == 'Girl'])
                #         girl['gender'] = 'Girl'
                #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                        
                #         boygirl = pd.concat([boy, girl], ignore_index=True)
                #         boygirl['before'] = boygirl['before'].apply(lambda x: round(x,2))
                #         boygirl['after'] = boygirl['after'].apply(lambda x: round(x,2))
                    
                #     elif q == 'TIME005':
                #         boy = time005(old_num.loc[old_num['gender'] == 'Boy'], new_num.loc[new_num['gender'] == 'Boy'])
                #         boy['gender'] = 'Boy'
                #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                        
                        
                #         girl = time005(old_num.loc[old_num['gender'] == 'Girl'], new_num.loc[new_num['gender'] == 'Girl'])
                #         girl['gender'] = 'Girl'
                #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                        
                #         boygirl = pd.concat([boy, girl], ignore_index=True)
                    
                #     else:
                #         if q in ['FVRT006A', 'FVRT006C']:
                #                 old_num, new_num = fvrt006(old_num, new_num, q)
                        
                #         # Count boy answers
                #         boy_before = old_num['answer'].loc[old_num['gender'] == 'Boy'].value_counts()
                #         boy_before = boy_before.apply(lambda x: x/len(old_denom.loc[old_denom['gender'] == 'Boy']))
                #         boy_after = new_num['answer'].loc[new_num['gender'] == 'Boy'].value_counts()
                #         boy_after = boy_after.apply(lambda x: x/len(new_denom.loc[new_denom['gender'] == 'Boy']))
                        
                #         # Convert the dataframes
                #         boy_before = pd.DataFrame({'name':boy_before.index, 'value':boy_before.values})
                #         boy_after = pd.DataFrame({'name':boy_after.index, 'value':boy_after.values})
                        
                #         # Merge before and after dataframes together
                #         boy = pd.merge(boy_before, boy_after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                        
                #         # Selecting where before is above zero to prevent dividebyzero error
                #         boy = boy.loc[boy['before'] > 0]
                        
                #         # Calculating growth
                #         boy['growth'] = boy.apply(lambda x: (x['after'] / 
                #          x['before'])-1, axis=1)
                        
                #         # Multiply by 100 to get growth percentages           
                #         boy['growth'] = boy['growth'].apply(lambda x: x*100)
                        
                        
                #         # Making a gender column so we can hue on it later
                #         boy['gender'] = 'Boy'
                        
                #         # Using the table_type function above
                #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                        
                        
                        
                #         # Count girl answers
                #         girl_before = old_num['answer'].loc[old_num['gender'] == 'Girl'].value_counts()
                #         girl_before = girl_before.apply(lambda x: x/len(old_denom.loc[old_denom['gender'] == 'Girl']))
                #         girl_after = new_num['answer'].loc[new_num['gender'] == 'Girl'].value_counts()
                #         girl_after = girl_after.apply(lambda x: x/len(new_denom.loc[new_denom['gender'] == 'Girl']))
                        
                #         # Convert to dataframes
                #         girl_before = pd.DataFrame({'name':girl_before.index, 'value':girl_before.values})
                #         girl_after = pd.DataFrame({'name':girl_after.index, 'value':girl_after.values})
                        
                #         # Merge before and after dataframes together
                #         girl = pd.merge(girl_before, girl_after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                        
                #         # Selecting where before is above zero to prevent dividebyzero error
                #         girl = girl.loc[girl['before'] > 0]
                        
                #         # Calculating growth
                #         girl['growth'] = girl.apply(lambda x: (x['after'] / 
                #         x['before'])-1, axis=1)
                        
                #         # Multiply by 100 to get growth percentages           
                #         girl['growth'] = girl['growth'].apply(lambda x: x*100)
                        
                        
                #         # Making a gender column so we can hue on it later
                #         girl['gender'] = 'Girl'
                        
                #         # Using the table_type function above
                #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                        
                #         # Concatenating the boy and girl tables into one dataframe
                #         boygirl = pd.concat([boy, girl], ignore_index=True)
                        
                #         # Multiply before and after by 100 and round to 2 decimal places
                #         boygirl['before'] = boygirl['before'].apply(lambda x: round(x*100,2))
                #         boygirl['after'] = boygirl['after'].apply(lambda x: round(x*100,2))
                
                # Appending old num and new num to the lists above
                old_nums.append(old_num)
                new_nums.append(new_num)
                #print(num['answer'])
                
                # Range to collect sample from
                s_num = num[olddate1:newdate2]
                
                # Sample size is the length of the s_num dataframe
                sample = len(s_num)
                
                # Append s_num to the s_nums list
                s_nums.append(s_num)
                
                # Getting total surveyed from old data
                sample_old = len(old_denom)
                
                
                # If total surveyed is zero, print invalid date range
                if sample_old == 0:
                    print('failed')
                    # sheet['{}'.format(loc)].offset(-1,0).value = ''
                    # sheet["{}".format(loc)].value = 'No data for this region!'
                else:
                    # If question code is one of the dodgy question codes, use the 
                    # functions for them above
                    if q == 'MONY011':
                        df_merge_col = money011(old_num, new_num)
                    elif q == 'TIME005':
                        df_merge_col = time005(old_num, new_num)
                    else:
                        if q in ['FVRT006A', 'FVRT006C']:
                            if b == False and g == False:
                                old_num, new_num = fvrt006(old_num, new_num, q)
                        
                        # Do value counts for old data
                        r_old = old_num['answer'].value_counts()
                        
                        # Getting percentage values that match the portal
                        r_old = r_old.apply(lambda x: x/sample_old)
                        
                        # Getting sample of new data
                        sample_new = len(new_denom)
                        
                        # getting value counts of new data
                        r_new = new_num['answer'].value_counts()
                        
                        # Getting percentage values that match the portal
                        r_new = r_new.apply(lambda x: x/sample_new)
                        
                        # Making dataframe to display values for each time period
                        before = pd.DataFrame({'name':r_old.index, 'value':r_old.values})
                        after = pd.DataFrame({'name':r_new.index, 'value':r_new.values})
                        
                        # Mergeing the dataframes ready for comparison
                        df_merge_col = pd.merge(before, after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                        df_merge_col = df_merge_col.loc[df_merge_col['before'] > 0]
                        print(r_new)
                        # Calculating growth
                        df_merge_col['growth'] = df_merge_col.apply(lambda x: (x['after'] / 
                                                                               x['before'])-1, axis=1)
                        
                        # Scaling by 100 and sorting from largest to smallest            
                        df_merge_col['growth'] = df_merge_col['growth'].apply(lambda x: x*100)
                        
                    # Sort values by filters the user has inputted
                    df_merge_col = df_merge_col.sort_values(by=sort, ascending=asc)
                                
                    # If question code is TIME005, get a copy of the dataframe
                    if q == 'TIME005':
                        # Displaying time in another format, making use of the hour_mins
                        # function above
                        time = df_merge_col.copy()
                        time['before'] = time['before'].apply(lambda x: "'{}'".format(str(datetime.timedelta(hours=x))))
                        time['before'] = time['before'].apply(lambda x: hour_mins(x))
                        time['after'] = time['after'].apply(lambda x: "'{}'".format(str(datetime.timedelta(hours=x))))
                        time['after'] = time['after'].apply(lambda x: hour_mins(x))
                        
                    # Creating dataframe in the style the user wants
                    final = table_type(df_merge_col, selnames, top, bottom, threshold, sign, sort)
                    
                    # If q is MONY011 round before and after columns to 2 decimal places 
                    if q == 'MONY011':
                        final['before'] = final['before'].apply(lambda x: round(x,2))
                        final['after'] = final['after'].apply(lambda x: round(x,2))
                    
                    # If code is TIME005 do nothing
                    elif q == 'TIME005':
                        final = final
                    else:
                        # Else multiply before and after columns by 100 and round to 
                        # 2 decimal places
                        final['before'] = final['before'].apply(lambda x: round(x*100,2))
                        final['after'] = final['after'].apply(lambda x: round(x*100,2))
                    
                    # Round growth column to nearest integer
                    final['growth'] = final['growth'].apply(lambda x: int(round(x)))
                    
                    # If products_tup isn't a tuple, replace the brackets with empty strings
                    if isinstance(products_tup, tuple) == False:
                        p_final = products_tup.replace("('", "")
                        p_final = p_final.replace("')", "")
                    else:
                        # Else make p_final equal to products_tup
                        p_final = products_tup
                    
                    
                    # Colour wheel
                    colors = ['#96D5EE', '#27AAE1', '#2D2E83', '#7277C7',
                      '#00A19A', '#ACCC00', '#662483', '#9D96C9',
                      '#CE2F6C', '#FAA932', '#464B9A', '#347575',
                      '#6AC2C2', '#76AF41', '#A33080', '#EB8EDB',
                      '#F66D9B', '#D95D27', '#F7915D', '#FFC814']  
                    
                    # # If display is True
                    # if display == True:
                    # If user has set graphs to Yes
                    if graph == 'Bar Chart':
                        # count starting at 1
                        c = 1
                        for col in ['before', 'after', 'growth']:
                            # if gender == "('Boy', 'Girl')":
                            #     try:
                            #         # Setting user inputted text colour
                            #         plt.rcParams['text.color'] = text_colour
                
                            #         # Vertical hue graph
                            #         if graph_orient == 'vertical':
                            #             graph = sns.catplot(x='name', y=col,
                            #             data=boygirl,
                            #             kind='bar', palette=['#D95D27', '#662483'],
                            #             hue='gender',
                            #             height=6, aspect=2, legend=False,
                            #             orient='v')
                            #         else:
                            #             # Horizontal hue graph
                            #             graph = sns.catplot(x=col, y='name',
                            #             data=boygirl,
                            #             kind='bar', palette=['#D95D27', '#662483'],
                            #             hue='gender',
                            #             height=6, aspect=2, legend=False,
                            #             orient='h')
                
                
                            #         # Fonts for title and labels
                            #         tfont = {'fontname':'{}'.format(title_font)}
                            #         lfont = {'fontname':'{}'.format(label_font)}
                
                            #         # Set user inputted title with correct formating
                            #         plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
                
                            #         # Adding correct formating to graph labels
                            #         if graph_orient == 'vertical':
                            #             plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                            #             plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                            #         else:
                            #             # Swapping axis labels for horizontal graph
                            #             plt.xlabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                            #             plt.ylabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour) 
                
                            #         # Rotate x-axis ticks based on what user has
                            #         # selected for tick_rotation
                            #         plt.xticks(rotation=tick_rotation, fontsize=ticksize)
                            #         plt.yticks(fontsize=ticksize)
                                    
                            #         ax = plt.gca()
                                    
                            #         if graph_orient == 'vertical':
                            #             ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                            #         else:
                            #             ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                    
                            #         # If halo is set to On
                            #         if halo == 'On':
                            #             if q == 'TIME005':
                            #                 if c != 3:
                            #                     # If count = 3 don't covert to hours and mins
                            #                     bar_labels(ax, time='y', orientation=graph_orient)
                            #                 else:
                            #                     bar_labels(ax, orientation=graph_orient)
                            #             else:
                            #                 bar_labels(ax, orientation=graph_orient)
                                    
                            #         # Setting colour of the tick labels
                            #         [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
                            #         [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
                            #         [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
                            #         [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
                                    
                            #         # Adding legend
                            #         plt.legend(bbox_to_anchor=(0.6, -0.1), ncol=2, frameon=False).set_title('')
                                    
                                    
                            #         # Despine graph
                            #         sns.despine(left=True)
                                    
                            #         # Getting current time for filenames
                            #         current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
                                
                            #     except:
                            #         print('failed')
                            #         # If there was an error with the hue graphs,
                            #         # print this message
                            #         # sheet["A38"].value = 'Could not create hue graph'
                
                            # Main graphs
                            try:
                                # Setting user inputted text colour
                                plt.rcParams['text.color'] = text_colour
                                
                                # Vertical graph
                                if graph_orient == 'vertical':
                                    graph = sns.catplot(x='name', y=col,
                                    data=final,
                                    kind='bar', palette=colors,
                                    height=6, aspect=2, legend=False,
                                    orient='v')
                                
                                # Horizontal graph
                                else:
                                    graph = sns.catplot(x=col, y='name',
                                    data=final,
                                    kind='bar', palette=colors,
                                    height=6, aspect=2, legend=False,
                                    orient='h')
                                
                                # Fonts for title and labels
                                tfont = {'fontname':'{}'.format(title_font)}
                                lfont = {'fontname':'{}'.format(label_font)}
                                
                                # Set user inputted title with correct formating
                                plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
                                
                                
                                # Adding correct formating to graph labels
                                if graph_orient == 'vertical':
                                    plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                    plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                else:
                                    # Swapping axis labels for horizontal graph
                                    plt.xlabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                    plt.ylabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour) 
                                
                                # Rotate x-axis ticks based on what user has
                                # selected for tick_rotation
                                plt.xticks(rotation=tick_rotation, fontsize=ticksize)
                                plt.yticks(fontsize=ticksize)
                                
                                ax = plt.gca()
                                
                                if graph_orient == 'vertical':
                                    ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                else:
                                    ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                
                                # If halo is set to On
                                if halo == 'On':
                                    if q == 'TIME005':
                                        if c != 3:
                                            # If count = 3 don't covert to hours and mins
                                            bar_labels(ax, time='y', orientation=graph_orient)
                                        else:
                                            bar_labels(ax, orientation=graph_orient)
                                    else:
                                        bar_labels(ax, orientation=graph_orient)
                                
                                
                                # Setting colour of the tick labels       
                                [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
                                [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
                                [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
                                [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
                                
                                # Despine graphs
                                sns.despine(left=True)
                                
                                # Getting current time for filenames
                                current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
                                
                                
                                # Add 1 to the count
                                c += 1
                            except:
                                print('failed')
                                # sheet["A38"].value = 'Invalid colour'
                
                    elif graph == 'Line Graph':
                        try:
                            line_graph(by_quarter, smooth, newdate2, num, denom, month_range, product, sumavg, selnames, text_colour, graph_title, title_font, title_fontsize, xaxis, yaxis, label_font, label_fontsize, tick_rotation, ticksize)
                        except:
                            print('failed')
                            # sheet["A38"].value = 'Could not create graph'
                    elif graph == 'Pie Chart':
                        try:
                            pie_chart(final, sumavg, product, text_colour, graph_title, title_fontsize, title_font, pie_label_col)
                        except:
                            print('failed')
                            # sheet["A38"].value = 'Could not create graph'
                
                
                    # If question code isn't in MONY011 or TIME005 then add % sign 
                    # to before and after columns
                    if q not in['MONY011', 'TIME005']:
                        final['before'] = final['before'].apply(lambda x: str(x) + '%')
                        final['after'] = final['after'].apply(lambda x: str(x) + '%')
                    
                    # Add % sign to the growth column
                    final['growth'] = final['growth'].apply(lambda x: str(x) + '%')
                    
                    
                    # If question code is TIME005 set table style and round the growth
                    # column to the nearest integer, add a % sign
                    if q == 'TIME005':
                        final = table_type(time, selnames, top, bottom, threshold, sign, sort)
                        final['growth'] = final['growth'].apply(lambda x: str(int(round(x))) + '%')
                    
                p = 0
                pretty_table(final, b, g, p, q, question_text)
                    
            # If user has selected Sum             
            else:
                st.write('Going through second sum loop')
                # For every product the user has inputted    
                for p in products_tup:
                    # Writing query for the numerator
                    query_num = """select prof.[record_id] as profile_record_id
                    ,ans.[record_id] as answer_record_id
                    ,[product]
                    ,[date_submitted]
                    ,[age]
                    ,[gender]
                    ,[location]
                    ,[question_code]
                    ,[answer]
                    ,[subquestion]
                    ,[survey_type]
                    from dat.t_profile as prof
                    LEFT JOIN  dat.t_answers  as ans ON (prof.record_id = ans.record_id and question_code='{0}')
                    WHERE prof.date_submitted between '2017/07/01' and '2030/10/31'
                    and gender in {1}
                    and age between {2} and {3}
                    and survey_type in {4}
                    and product = {5} """.format(q if q not in ['FVRT006A', 'FVRT006C'] else 'FVRT006', gender, age1,
                                                 age2, survey,
                                                 product_string if len(products_tup) == 1 else p)
                    
    
                    # Writing query for the denominator
                    query_denom = """SELECT [record_id]
                    , [date_submitted]
                    , [product]
                    ,[survey_type]
                    ,[gender]
                    FROM [dat].[t_profile]
                    WHERE gender in {0}
                    and age between {1} and {2}
                    and date_submitted between '2017/07/01' and '2030/10/31'
                    and survey_type in {3}
                    and product = {4} """.format(gender, age1, age2, survey,
                                                 product_string if len(products_tup) == 1 else p)
    
                    # Getting correct question text for inputted question code
                    text_query = """select top(1) [question_code]
                    ,[question_text]
                    from [dat].[t_question_codes]
                    where question_code = '{}' """.format(q if q not in ['FVRT006A', 'FVRT006C'] else 'FVRT006')
                    
                    Query_text = pd.read_sql(text_query, cnxn)
                    text = pd.DataFrame(Query_text)
                    question_text = text['question_text'].values[0]
                      
                    # Putting numerator into a dataframe
                    Query_num = pd.read_sql(query_num, cnxn)
                    num = pd.DataFrame(Query_num)
                    num['date'] = pd.to_datetime(num['date_submitted'])
                    num = num.set_index(num['date'])
                    num = num.sort_index()
                    num = num.replace(['None','2BCLEANED'], np.nan)
                     
                    
                    # Putting denominator into a dataframe
                    Query_denom = pd.read_sql(query_denom, cnxn)
                    denom = pd.DataFrame(Query_denom)
                    denom['date'] = pd.to_datetime(denom['date_submitted'])
                    denom = denom.set_index(denom['date'])
                    denom = denom.sort_index()
                    denom = denom.replace(['None','2BCLEANED'], np.nan)
                    
                    
                    # Selecting dates within the dataframe
                    old_num = num[olddate1:olddate2]
                    old_denom = denom[olddate1:olddate2]
                    new_num = num[newdate1:newdate2]
                    new_denom = denom[newdate1:newdate2]
                    
                    # If there is a subquestion, pass it through the subquestions function
                    if subq != '':
                        st.write("Shouldn't be going through this loop")
                        old_num, new_num = subquestions(old_num, new_num, subq)
                    
                    # # If gender is Boy/Girl, set up conditions for the hued gender graph
                    # if gender == "('Boy', 'Girl')":
                    #     # Need to manipulate dataframes based on question code the user has
                    #     # selected
                    #     if q == 'MONY011':
                    #         boy = money011(old_num.loc[old_num['gender'] == 'Boy'], new_num.loc[new_num['gender'] == 'Boy'])
                    #         # boy = boy.sort_values(by=sort, ascending=asc)
                    #         boy['gender'] = 'Boy'
                    #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                            
                            
                    #         girl = money011(old_num.loc[old_num['gender'] == 'Girl'], new_num.loc[new_num['gender'] == 'Girl'])
                    #         # girl = girl.sort_values(by=sort, ascending=asc)
                    #         girl['gender'] = 'Girl'
                    #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                            
                    #         boygirl = pd.concat([boy, girl], ignore_index=True)
                    #         boygirl['before'] = boygirl['before'].apply(lambda x: round(x,2))
                    #         boygirl['after'] = boygirl['after'].apply(lambda x: round(x,2))
                        
                    #     elif q == 'TIME005':
                    #         boy = time005(old_num.loc[old_num['gender'] == 'Boy'], new_num.loc[new_num['gender'] == 'Boy'])
                    #         # boy = boy.sort_values(by=sort, ascending=asc)
                    #         boy['gender'] = 'Boy'
                    #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                            
                            
                    #         girl = time005(old_num.loc[old_num['gender'] == 'Girl'], new_num.loc[new_num['gender'] == 'Girl'])
                    #         # girl = girl.sort_values(by=sort, ascending=asc)
                    #         girl['gender'] = 'Girl'
                    #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                            
                    #         boygirl = pd.concat([boy, girl], ignore_index=True)
                        
                    #     else:
                    #         if q in ['FVRT006A', 'FVRT006C']:
                    #             old_num, new_num = fvrt006(old_num, new_num, q)
                            
                            
                    #         # Count boy answers    
                    #         boy_before = old_num['answer'].loc[old_num['gender'] == 'Boy'].value_counts()
                    #         boy_before = boy_before.apply(lambda x: x/len(old_denom.loc[old_denom['gender'] == 'Boy']))
                    #         boy_after = new_num['answer'].loc[new_num['gender'] == 'Boy'].value_counts()
                    #         boy_after = boy_after.apply(lambda x: x/len(new_denom.loc[new_denom['gender'] == 'Boy']))
                            
                    #         # Convert to dataframes
                    #         boy_before = pd.DataFrame({'name':boy_before.index, 'value':boy_before.values})
                    #         boy_after = pd.DataFrame({'name':boy_after.index, 'value':boy_after.values})
                            
                    #         # Merge dataframes
                    #         boy = pd.merge(boy_before, boy_after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                    #         boy = boy.loc[boy['before'] > 0]
                            
                    #         # Calculating growth
                    #         boy['growth'] = boy.apply(lambda x: (x['after'] / 
                    #           x['before'])-1, axis=1)
                            
                    #         # Multiplying growth column by 100           
                    #         boy['growth'] = boy['growth'].apply(lambda x: x*100)
                            
                    #         # # Sort values based on filters the user has selected
                    #         # boy = boy.sort_values(by=sort, ascending=asc)
                            
                    #         # Setting a gender column equal to Boy
                    #         boy['gender'] = 'Boy'
                            
                    #         # Using the table_type function above
                    #         boy = table_type(boy, selnames, top, bottom, threshold, sign, sort)
                            
                            
                            
                    #         # Count girl answers
                    #         girl_before = old_num['answer'].loc[old_num['gender'] == 'Girl'].value_counts()
                    #         girl_before = girl_before.apply(lambda x: x/len(old_denom.loc[old_denom['gender'] == 'Girl']))
                    #         girl_after = new_num['answer'].loc[new_num['gender'] == 'Girl'].value_counts()
                    #         girl_after = girl_after.apply(lambda x: x/len(new_denom.loc[new_denom['gender'] == 'Girl']))
                            
                    #         # Convert the dataframes
                    #         girl_before = pd.DataFrame({'name':girl_before.index, 'value':girl_before.values})
                    #         girl_after = pd.DataFrame({'name':girl_after.index, 'value':girl_after.values})
                            
                    #         # Merge dataframes
                    #         girl = pd.merge(girl_before, girl_after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                    #         girl = girl.loc[girl['before'] > 0]
                            
                    #         # Calculating growth
                    #         girl['growth'] = girl.apply(lambda x: (x['after'] / 
                    #         x['before'])-1, axis=1)
                            
                    #         # Multiplying growth column by 100            
                    #         girl['growth'] = girl['growth'].apply(lambda x: x*100)
                             
                    #         # # Sort values based on filters the user has selected 
                    #         # girl = girl.sort_values(by=sort, ascending=asc)
                            
                    #         # Setting a gender column equal to Girl
                    #         girl['gender'] = 'Girl'
                            
                    #         # Using the table_type function above
                    #         girl = table_type(girl, selnames, top, bottom, threshold, sign, sort)
                            
                    #         # Concatenating the boy and girl tables into one dataframe
                    #         boygirl = pd.concat([boy, girl], ignore_index=True)
                            
                    #         # Multiply before and after by 100 and round to 2 decimal places
                    #         boygirl['before'] = boygirl['before'].apply(lambda x: round(x*100,2))
                    #         boygirl['after'] = boygirl['after'].apply(lambda x: round(x*100,2))
                    
                    
                    # Appending old num and new num to the lists above
                    old_nums.append(old_num)
                    new_nums.append(new_num)
                    
                    # Range to collect sample from
                    s_num = num[olddate1:newdate2]
                    
                    # Sample size is the length of the s_num dataframe
                    sample = len(s_num)
                    
                    # Append s_num to the s_nums list
                    s_nums.append(s_num)
                    
                    # Getting total surveyed from old data
                    sample_old = len(old_denom)
                    
                    
                    # If total surveyed is zero, print invalid date range
                    if sample_old == 0:
                        print('failed')
                        # sheet['{}'.format(loc)].offset(-1,0).value = ''
                        # sheet["{}".format(loc)].value = 'No data for this region!'
                    else:
                        # If question code is one of the dodgy question codes, use the 
                        # functions for them above
                        if q == 'MONY011':
                            df_merge_col = money011(old_num, new_num)
                        elif q == 'TIME005':
                            df_merge_col = time005(old_num, new_num)
                        else:
                            if q in ['FVRT006A', 'FVRT006C']:
                                if b == False and g == False:
                                    old_num, new_num = fvrt006(old_num, new_num, q)
                                    
                            # Do value counts for old data
                            r_old = old_num['answer'].value_counts()
                            
                            # Getting percentage values that match the portal
                            r_old = r_old.apply(lambda x: x/sample_old)
                            
                            # Getting sample of new data
                            sample_new = len(new_denom)
                            
                            # Do value counts for new data
                            r_new = new_num['answer'].value_counts()
         
                            # Getting percentage values that match the portal
                            r_new = r_new.apply(lambda x: x/sample_new)
                            
                            # Making dataframe to display values for each time period
                            before = pd.DataFrame({'name':r_old.index, 'value':r_old.values})
                            after = pd.DataFrame({'name':r_new.index, 'value':r_new.values})
                            # print(old_num)
                            # sleep(100000)
                            # Mergeing the dataframes ready for comparison
                            df_merge_col = pd.merge(before, after, on='name').rename(columns={'value_x': 'before', 'value_y': 'after'})
                            df_merge_col = df_merge_col.loc[df_merge_col['before'] > 0]
                            
                            # Calculating growth
                            df_merge_col['growth'] = df_merge_col.apply(lambda x: (x['after'] / 
                                                                                    x['before'])-1, axis=1)
                            
                            # Scaling growth by 100            
                            df_merge_col['growth'] = df_merge_col['growth'].apply(lambda x: x*100)
                        
                        # Sort values by filters the user has inputted
                        df_merge_col = df_merge_col.sort_values(by=sort, ascending=asc)
                        
                        # If question code is TIME005, get a copy of the dataframe
                        if q == 'TIME005':
                            # Displaying time in another format, making use of the hour_mins
                            # function above
                            time = df_merge_col.copy()
                            time['before'] = time['before'].apply(lambda x: "'{}'".format(str(datetime.timedelta(hours=x))))
                            time['before'] = time['before'].apply(lambda x: hour_mins(x))
                            time['after'] = time['after'].apply(lambda x: "'{}'".format(str(datetime.timedelta(hours=x))))
                            time['after'] = time['after'].apply(lambda x: hour_mins(x))
                          
                        # Creating dataframe in the style the user wants
                        final = table_type(df_merge_col, selnames, top, bottom, threshold, sign, sort)
                        
                        # If q is MONY011 round before and after columns to 2 decimal places
                        if q == 'MONY011':
                            final['before'] = final['before'].apply(lambda x: round(x,2))
                            final['after'] = final['after'].apply(lambda x: round(x,2))
                        
                        # If code is TIME005 do nothing
                        elif q == 'TIME005':
                            final = final
                        else:
                            # Else multiply before and after columns by 100 and round to 
                            # 2 decimal places
                            final['before'] = final['before'].apply(lambda x: round(x*100,2))
                            final['after'] = final['after'].apply(lambda x: round(x*100,2))
                        
                        # Round growth column to nearest integer
                        final['growth'] = final['growth'].apply(lambda x: int(round(x)))
                        
                        # Replace apostrophies in p with an empty string
                        p_final = p.replace("'", "")
                        
                        # Colour wheel
                        colors = ['#96D5EE', '#27AAE1', '#2D2E83', '#7277C7',
                                  '#00A19A', '#ACCC00', '#662483', '#9D96C9',
                                  '#CE2F6C', '#FAA932', '#464B9A', '#347575',
                                  '#6AC2C2', '#76AF41', '#A33080', '#EB8EDB',
                                  '#F66D9B', '#D95D27', '#F7915D', '#FFC814']  
                        
                        # # If display is True
                        # if display == True:
                        # If user has set graphs to Yes
                        if graph == 'Bar Chart':
                            # Count starting at 1
                            c = 1
                            for col in ['before', 'after', 'growth']:
                                # if gender == "('Boy', 'Girl')":
                                #     try:
                                #         # Setting user inputted text colour
                                #         plt.rcParams['text.color'] = text_colour
                                        
                                #         # Vertical graph
                                #         if graph_orient == 'vertical':
                                #             graph = sns.catplot(x='name', y=col,
                                #             data=boygirl,
                                #             kind='bar', palette=['#D95D27', '#662483'],
                                #             hue='gender',
                                #             height=6, aspect=2, legend=False,
                                #             orient='v')
                                #         else:
                                #             # Horizontal graph
                                #             graph = sns.catplot(x=col, y='name',
                                #             data=boygirl,
                                #             kind='bar', palette=['#D95D27', '#662483'],
                                #             hue='gender',
                                #             height=6, aspect=2, legend=False,
                                #             orient='h')
                                        
                                        
                                #         # Fonts for title and labels
                                #         tfont = {'fontname':'{}'.format(title_font)}
                                #         lfont = {'fontname':'{}'.format(label_font)}
                                        
                                #         # Set user inputted title with correct formating
                                #         plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
                                        
                                #         # Adding correct formatting from graph labels
                                #         if graph_orient == 'vertical':
                                #             plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                #             plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                #         else:
                                #             # Swapping axis labels for horizontal graph
                                #             plt.xlabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                #             plt.ylabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour) 
                                        
                                #         # Rotating x-axis ticks based on what the user has selected
                                #         plt.xticks(rotation=tick_rotation, fontsize=ticksize)
                                #         plt.yticks(fontsize=ticksize)
                                        
                                #         ax = plt.gca()
                                        
                                #         if graph_orient == 'vertical':
                                #             ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                #         else:
                                #             ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                        
                                #         if halo == 'On':
                                #             # If halo is set to On
                                #             if q == 'TIME005':
                                #                 if c != 3:
                                #                     # If count = 3, don't convert to hours and mins
                                #                     bar_labels(ax, time='y', orientation=graph_orient)
                                #                 else:
                                #                     bar_labels(ax, orientation=graph_orient)
                                #             else:
                                #                 bar_labels(ax, orientation=graph_orient)
                                        
                                #         # Setting colour of the tick labels
                                #         [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
                                #         [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
                                #         [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
                                #         [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
                                        
                                #         # Plotting a legend
                                #         plt.legend(bbox_to_anchor=(0.6, -0.1), ncol=2, frameon=False).set_title('')
                                        
                                #         # Despine graphs
                                #         sns.despine(left=True)
                                        
                                #         # Getting current time for filenames
                                #         current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
                                        
                                #     except:
                                #         print('failed')
                                #         # If there was an error with the hue graphs,
                                #         # print this message
                                #         # sheet["A38"].value = 'Could not create hue graph'
                    
                                try:
                                #Setting user inputted text colour
                                    plt.rcParams['text.color'] = text_colour
                                    
                                    # Vertical graph
                                    if graph_orient == 'vertical':
                                        graph = sns.catplot(x='name', y=col,
                                        data=final,
                                        kind='bar', palette=colors,
                                        height=6, aspect=2, legend=False,
                                        orient='v')
                                    else:
                                        # Horizontal graph
                                        graph = sns.catplot(x=col, y='name',
                                        data=final,
                                        kind='bar', palette=colors,
                                        height=6, aspect=2, legend=False,
                                        orient='h')
                                        
    
                                    # Fonts for title and labels
                                    tfont = {'fontname':'{}'.format(title_font)}
                                    lfont = {'fontname':'{}'.format(label_font)}
                                    
                                    # Set user inputted title with correct formating
                                    plt.title('{}'.format(graph_title),**tfont, fontsize=title_fontsize)
                                    
                                    # Adding correct formatting from graph labels
                                    if graph_orient == 'vertical':
                                        plt.ylabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                        plt.xlabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                    else:
                                        # Swapping axis labels for horizontal graph
                                        plt.xlabel('{}'.format(yaxis),**lfont, fontsize=label_fontsize, color=text_colour)
                                        plt.ylabel('{}'.format(xaxis),**lfont, fontsize=label_fontsize, color=text_colour) 
                                    
                                    # Rotating x-axis ticks based on what the user has selected
                                    plt.xticks(rotation=tick_rotation, fontsize=ticksize)
                                    plt.yticks(fontsize=ticksize)
                                    
                                    ax = plt.gca()
                                    if graph_orient == 'vertical':
                                        ax.yaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                    else:
                                        ax.xaxis.set_major_formatter(mtick.PercentFormatter(decimals=0))
                                    
                                    if halo == 'On':
                                        # If halo is set to On
                                        if q == 'TIME005':
                                            if c != 3:
                                                # If count = 3, don't convert to hours and mins
                                                bar_labels(ax, time='y', orientation=graph_orient)
                                            else:
                                                bar_labels(ax, orientation=graph_orient)
                                        else:
                                            bar_labels(ax, orientation=graph_orient)
                                    
                                    # Setting colour of the tick labels        
                                    [t.set_color(text_colour) for t in ax.xaxis.get_ticklines()]
                                    [t.set_color(text_colour) for t in ax.xaxis.get_ticklabels()]
                                    [t.set_color(text_colour) for t in ax.yaxis.get_ticklines()]
                                    [t.set_color(text_colour) for t in ax.yaxis.get_ticklabels()]
                                    
                                    # Despine graphs
                                    sns.despine(left=True)
                                    
                                    # Getting current time for filenames
                                    current_time = str(datetime.datetime.now().strftime("%H_%M_%S"))
    
                                    # Save graphs
                                    # graph.savefig('{0}{1}_{2}_{3}.png'.format(filepath, p_final, col, current_time), transparent=True)
                                    # Add 1 to count
                                    c += 1
                                except:
                                    print('failed')
                                    # sheet["A38"].value = 'Invalid colour'
                                # except Exception as e: print(e)
                                # sleep(100000)
                    
                        elif graph == 'Line Graph':
                            try:
                                line_graph(by_quarter, smooth, newdate2, num, denom, month_range, product, sumavg, selnames, text_colour, graph_title, title_font, title_fontsize, xaxis, yaxis, label_font, label_fontsize, tick_rotation, ticksize)
                            except:
                                print('failed')
                                # sheet["A38"].value = 'Could not create graph'
                        elif graph == 'Pie Chart':
                            try:
                                pie_chart(final, sumavg, product, text_colour, graph_title, title_fontsize, title_font, pie_label_col)
                            except:
                                print('failed')
                                # sheet["A38"].value = 'Could not create graph'
                        
                        # If question code isn't in MONY011 or TIME005 then add % sign 
                        # to before and after columns
                        if q not in['MONY011', 'TIME005']:
                            final['before'] = final['before'].apply(lambda x: str(x) + '%')
                            final['after'] = final['after'].apply(lambda x: str(x) + '%')
                        
                        # Add % sign to the growth column
                        final['growth'] = final['growth'].apply(lambda x: str(x) + '%')
                        
                        # If question code is TIME005 set table style and round the growth
                        # column to the nearest integer, add a % sign
                        if q == 'TIME005':
                            final = table_type(time, selnames, top, bottom, threshold, sign, sort)
                            final['growth'] = final['growth'].apply(lambda x: str(int(round(x))) + '%')
    
                    pretty_table(final, b, g, p, q, question_text)

           
                        


olddate1, olddate2, newdate1, newdate2, age1, age2, b, g = 0, 0, 0, 0, 0, 0, 0, 0
product, sumavg, qcode, subq, offline, digital, selnames, top = 0, 0, 0, 0, 0, 0, 0, 0
bottom, sign, threshold, sort, asc, graph, graph_orient, text_colour = 0, 0, 0, 0, 0, 0, 0, 0
month_range, pie_label_col, by_quarter, smooth, ticksize = 0, 0, 0, 0, 0
graph_title, title_font, title_fontsize, xaxis, yaxis, label_font = 0, 0, 0, 0, 0, 0
label_fontsize, tick_rotation, halo, answer, comparison = 0, 0, 0, 0, 0
sub_comparison = 0




TIF_logo = Image.open('Y:/Python/Streamlit/TIF_logo.png')
st.image(TIF_logo)
st.title('Universal Tool')
#---------------------------------Filters-------------------------------------
# STEP 1

# Header
st.sidebar.header('Step 1: Select some basic filters')
st.sidebar.text('')

# Date filters
olddate1 = st.sidebar.date_input('Date Ranges', datetime.date(2021,1,1), key='1')
olddate2 = st.sidebar.date_input('', datetime.date.today(), key='2')

st.sidebar.text('')

newdate1 = st.sidebar.date_input('Comparison Date Ranges',
                                 datetime.date(2021,1,1), key='3')
newdate2 = st.sidebar.date_input('', datetime.date.today(), key='4')

st.sidebar.text('')

# Gender Filter
st.sidebar.text('Gender')
                      
b = st.sidebar.checkbox('Boy')
g = st.sidebar.checkbox('Girl')

st.sidebar.text('')

# Age Filter
ages = st.sidebar.slider('Select Age Range', 3, 18, (6, 12))
age1 = list(ages)[0]
age2 = list(ages)[1]
st.sidebar.text('')

# Question Code Filters
qcode = st.sidebar.text_input('Question Code(s):', key='1')
st.sidebar.text('')
subq = st.sidebar.text_input('Subquestion(s):', key='2')
st.sidebar.text('')

st.sidebar.text('Survey Type')
offline = st.sidebar.checkbox('offline')
digital = st.sidebar.checkbox('digital')

st.sidebar.text('')
product = st.sidebar.text_input('Product(s):')
st.sidebar.text('')

sumavg = st.sidebar.radio('Sum or Average?',
                 ('Sum', 'Average'))

st.sidebar.text('')



# STEP 2

# Header
st.sidebar.header('Step 2: Select your table filters')
st.sidebar.text('')



# Create Top Filter
top = st.sidebar.slider('Show Top', 0, 50, 10)

# If top filter is set to zero, display bottom filter in the sidebar
if top == 0:
    # Create Bottom Filter
    bottom = st.sidebar.slider('Show Bottom', 0, 50, 0)

    st.sidebar.text('')

    # If top and bottom filters are both set to zero,
    # display threshold filters
    if (top == 0) & (bottom == 0):
        # Create Select Names Filter
        selnames = st.sidebar.text_input('Select Name(s):')
        st.sidebar.text('')
        
        if selnames == '':
            # Create Threshold Filters
            sign = st.sidebar.radio('Threshold Sign:',
                             ('>', '<', '='))
            st.sidebar.text('')
    
            threshold = st.sidebar.slider('Threshold %:', 0, 50, 0)
            
sort = st.sidebar.radio('Sort Tables By:',
                 ('before', 'after', 'growth'))

asc = st.sidebar.radio('Ascending Order?:',
                 ('Yes', 'No'))
        
st.sidebar.text('')
st.sidebar.text('')


# STEP 3

# Header
st.sidebar.header('Step 3: Select your graph filters')
st.sidebar.text('')

# Select Graph Filter
graph = st.sidebar.radio('Graphs?',
                 ('Bar Chart', 'Line Graph', 'Pie Chart', 'None'))
st.sidebar.text('')

# If Select Graph Filter is not None, display other filters
if graph != 'None':
    # Graph Orientation Filter
    graph_orient = st.sidebar.radio('Graph Orientation',
                     ('Vertical', 'Horizontal'))
    st.sidebar.text('')

    # Text Colour Filter
    text_colour = st.sidebar.color_picker('Text Colour ', '#00f900')
    st.sidebar.text('')

    if graph == 'Line Graph':
        # Months Filter
        month_range = st.sidebar.slider('Months (Line Graph)', 1, 24, 12)
        st.sidebar.text('')

    if graph == 'Pie Chart':
        # Label Colour Filter
        pie_label_col = st.sidebar.color_picker('Label Colour (Pie Chart)', '#00f900')
        st.sidebar.text('')

    if graph == 'Line Graph':
        # Quarter Filter
        by_quarter = st.sidebar.radio('Quarters? (Line Graph)',
                         ('Yes', 'No'))
        st.sidebar.text('')

        # Smooth Lines Filter
        smooth = st.sidebar.radio('Smooth Line? (Line Graph)',
                         ('Vertical', 'Horizontal'))
        st.sidebar.text('')

    # General Graph Labelling Filters
    ticksize = st.sidebar.slider('Tick Fontsize', 1, 36, 10)
    st.sidebar.text('')

    graph_title = st.sidebar.text_input('Graph Title:', key='1')
    st.sidebar.text('')
    title_font = st.sidebar.text_input('Title Font:', key='2')
    st.sidebar.text('')

    title_fontsize = st.sidebar.slider('Title Fontsize', 1, 36, 20)
    st.sidebar.text('')


    xaxis = st.sidebar.text_input('X-Axis Label:', key='3')
    st.sidebar.text('')
    yaxis = st.sidebar.text_input('Y-Axis Label:', key='4')
    st.sidebar.text('')
    label_font = st.sidebar.text_input('Label Font:', key='5')
    st.sidebar.text('')

    label_fontsize = st.sidebar.slider('Label Fontsize', 1, 36, 14)
    st.sidebar.text('')

    tick_rotation = st.sidebar.radio('Tick Rotation',
                     ('Vertical', 'Horizontal'))
    st.sidebar.text('')

    halo = st.sidebar.radio('Halos',
                     ('On', 'Off'))
    st.sidebar.text('')

    st.sidebar.text('')


# STEP 4

# Header
st.sidebar.header('Step 4 (optional): Select your comparison filters')
st.sidebar.text('')

# Comparison Filters
answer = st.sidebar.text_input('Answer:', key='6')
st.sidebar.text('')
comparison = st.sidebar.text_input('Comparison:', key='7')
st.sidebar.text('')
sub_comparison = st.sidebar.text_input('Sub-Comparison:', key='8')
st.sidebar.text('')

st.write('')

button1, button2, button3 = st.beta_columns(3)
with button1:
    if st.button('Press me!', key='1'):
        st.write('')
        st.write('')
        st.write('')
        get_data = get_data(olddate1, olddate2, newdate1, newdate2, age1, age2, b, g,
                            product, sumavg, qcode, subq, offline, digital, selnames, top,
                            bottom, sign, threshold, sort, asc, graph, graph_orient, text_colour,
                            month_range, pie_label_col, by_quarter, smooth, ticksize,
                            graph_title, title_font, title_fontsize, xaxis, yaxis, label_font,
                            label_fontsize, tick_rotation, halo, answer, comparison,
                            sub_comparison)
        # st.dataframe(get_data)
with button2:
    st.button('Press me!', key='2')
    
with button3:
    st.button('Press me!', key='3')