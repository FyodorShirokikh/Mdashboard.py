########################################################################################################################
#
#   Making simple Dashboard to show progress of the Monitoring plan
#   v.1.0
#   Feb 01, 2023
#
#   Input: \\BCCFS-HQ\reserve_data\отчеты УКР ДКР\Филиалы\1. Мониторинг\4. Мониторинг 2023\
#                                                                       График планового мониторинга 2023.xlsx,
#
#          \\BCCFS-HQ\reserve_data\отчеты УКР ДКР\ГО\Заявки на согласование ОМОД\Новый журнал по заявкам 2023.xlsm
#
#   Output: Dashboard
#   IP & Port: localhost:8000
#
########################################################################################################################
import numpy as np
import pandas as pd
import dash
from datetime import date, datetime
from dash import dcc, html, dash_table, ctx
import plotly.express as px
from dash.dependencies import Input, Output
import plotly.graph_objects as go
import dash_bootstrap_components as dbc
import warnings

#   Silent mode
warnings.simplefilter(action='ignore', category=UserWarning)
pd.options.mode.chained_assignment = None  # default='warn'


# ---------------------------------------------------------------------
#   Loading file content of Новый журнал по заявкам 2023.xlsm file
# ---------------------------------------------------------------------
def load_new_jornal(filename):
    df = pd.DataFrame()
    try:
        df = pd.read_excel(filename, sheet_name='Лист1', skiprows=6)
    except FileNotFoundError:
        print("File Новый журнал по заявкам 2023.xlsm does not exist")
    return df


# ---------------------------------------------------------------------
#   Loading file content of График планового мониторинга 2023.xlsx file
# ---------------------------------------------------------------------
def load_mon(filename):
    df = pd.DataFrame()
    try:
        df = pd.read_excel(filename, sheet_name='График', skiprows=0)
    except FileNotFoundError:
        print("График планового мониторинга 2023.xlsx")
    return df


# ---------------------------------------------------------------------
#   # Find the leaders
# ---------------------------------------------------------------------
def top_leaders(df_mon_f, date_start, date_end):
    df_top_f = pd.DataFrame()
    try:
        #   Changing type of some fields (from Object to Date)
        df_mon_f['Дата заключения УМОД'] = df_mon_f['Дата заключения УМОД'].astype("datetime64[ns]")

        #   Reduction the range by date
        df_top_f = df_mon_f[(df_mon_f['Дата заключения УМОД'] >= date_start + ' 00:00:00') &
                            (df_mon_f['Дата заключения УМОД'] <= date_end + ' 23:59:59')]

        #############################################################################################################
        #
        # Count for each employee (Choose field "Исполнитель" if you use "Новый журнал по заявкам 2023.xlsm" or
        #                               field "Ф.И.О. исполнителя" if you use "График планового мониторинга 2023.xlsx")
        #
        #############################################################################################################
        # df_top_f = df_top_f[["Исполнитель", "Дата заключения УМОД"]] \
        #     .groupby(['Исполнитель']) \
        #     .count().reset_index()
        df_top_f = df_top_f[["Ф.И.О. исполнителя", "Дата заключения УМОД"]] \
            .groupby(['Ф.И.О. исполнителя']) \
            .count().reset_index()

        # Rename column
        df_top_f.rename(columns={'Дата заключения УМОД': 'Количество'}, inplace=True)

        # First 20 employees
        df_top_f = df_top_f.sort_values(by=["Количество", "Ф.И.О. исполнителя"],
                                        ascending=[False, True]).groupby(['Ф.И.О. исполнителя'], as_index=False).nth[
                   :20]
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_top_f


# ---------------------------------------------------------------------
#   # Top of branches
# ---------------------------------------------------------------------
def top_branches(df_mon_f, date_start, date_end):
    df_top_f = pd.DataFrame()
    try:
        #   Changing type of some fields (from Object to Date)
        df_mon_f['Дата заключения УМОД'] = df_mon_f['Дата заключения УМОД'].astype("datetime64[ns]")

        #   Reduction the range of date
        df_top_f = df_mon_f[(df_mon_f['Дата заключения УМОД'] >= date_start + ' 00:00:00') &
                            (df_mon_f['Дата заключения УМОД'] <= date_end + ' 23:59:59')]

        # Count for each branch
        df_top_f = df_top_f[["Филиал Банка (рассмотрения заявки)", "Дата заключения УМОД"]] \
            .groupby(['Филиал Банка (рассмотрения заявки)']) \
            .count().reset_index()

        # Rename column
        df_top_f.rename(columns={'Дата заключения УМОД': 'Количество'}, inplace=True)

        # First branches
        df_top_f = df_top_f.sort_values(by=["Филиал Банка (рассмотрения заявки)", "Количество"],
                                        ascending=[True, False]).groupby(['Филиал Банка (рассмотрения заявки)'],
                                                                         as_index=False).nth[:10]
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_top_f


# ---------------------------------------------------------------------
#   # Find the anti-leaders from each branch
# ---------------------------------------------------------------------
def antitop_filials_r(df_mon_f, date_start, date_end):
    df_antitop_f = pd.DataFrame()
    try:
        #   Changing type of some fields (from Object to Date)
        df_mon_f['Дата заключения УМОД'] = df_mon_f['Дата заключения УМОД'].astype("datetime64[ns]")

        #   Reduction the range of date
        df_antitop_f = df_mon_f[(df_mon_f['Дата заключения УМОД'] >= date_start + ' 00:00:00') &
                                (df_mon_f['Дата заключения УМОД'] <= date_end + ' 23:59:59')]

        # Count for each employee
        df_antitop_f = df_antitop_f[["Филиал Банка (рассмотрения заявки)", "Исполнитель", "Дата заключения УМОД"]] \
            .groupby(['Филиал Банка (рассмотрения заявки)', 'Исполнитель']) \
            .count().reset_index()

        # Rename column
        df_antitop_f.rename(columns={'Дата заключения УМОД': 'Количество'}, inplace=True)

        # Anti-top employees
        df_antitop_f = df_antitop_f.sort_values(by=["Количество", "Филиал Банка (рассмотрения заявки)", "Исполнитель"],
                                                ascending=[True, True, True]).groupby(
            ['Филиал Банка (рассмотрения заявки)'],
            as_index=False).nth[:1]
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_antitop_f


# ---------------------------------------------------------------------
#   Calculate average time of monitoring
# ---------------------------------------------------------------------
def average_filials_r(df_temp_f, date_start, date_end):
    df_top_f = pd.DataFrame()
    try:
        #   Clean dataframe
        df_temp_f.dropna()

        #   Changing type of some fields (from Object to Date)
        df_temp_f['Дата заключения УМОД'] = df_temp_f['Дата заключения УМОД'].astype("datetime64[ns]")

        #   Reduction the range of date
        df_top_f = df_temp_f[(df_temp_f['Дата заключения УМОД'] >= date_start + ' 00:00:00') &
                             (df_temp_f['Дата заключения УМОД'] <= date_end + ' 23:59:59')]
        df_top_f = df_top_f[["Филиал исполнитель ДМОД", "Время затраченное на заключение Исполнителем"]]

        # Formatting to datetime
        df_top_f['Время затраченное на заключение Исполнителем'] = pd.to_datetime(
            df_top_f['Время затраченное на заключение Исполнителем'], format='%H:%M:%S')

        # Calc average time
        df_top_f = df_top_f.groupby('Филиал исполнитель ДМОД').mean().reset_index()

        # Formatting to string
        df_top_f['Время затраченное на заключение Исполнителем'] = df_top_f[
            'Время затраченное на заключение Исполнителем'].dt.strftime('%H:%M:%S')
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_top_f


# ---------------------------------------------------------------------
#   Calculating average time of monitoring, according categories
# ---------------------------------------------------------------------
def average_filials_unique(df_temp_f, date_start, date_end):
    df_top_f = pd.DataFrame()
    try:
        #   Clean
        df_temp_f.dropna()

        #   Changing type of some fields (from Object to Date)
        df_temp_f['Дата заключения УМОД'] = df_temp_f['Дата заключения УМОД'].astype("datetime64[ns]")

        #   Reduction the range of date
        df_top_f = df_temp_f[(df_temp_f['Дата заключения УМОД'] >= date_start + ' 00:00:00') &
                             (df_temp_f['Дата заключения УМОД'] <= date_end + ' 23:59:59')]
        df_top_f = df_top_f[["Уникальность обеспечения", "Филиал исполнитель ДМОД",
                             "Время затраченное на заключение Исполнителем"]]

        # Formatting to datetime
        df_top_f['Время затраченное на заключение Исполнителем'] = pd.to_datetime(
            df_top_f['Время затраченное на заключение Исполнителем'], format='%H:%M:%S')

        # Calculate average time
        df_top_f = df_top_f.groupby(['Уникальность обеспечения', 'Филиал исполнитель ДМОД']) \
            .agg({'Время затраченное на заключение Исполнителем': ['mean']}).reset_index()

        # Formatting to string
        df_top_f.columns = df_top_f.columns.map('_'.join)
        df_top_f['Время затраченное на заключение Исполнителем_mean'] = df_top_f[
            'Время затраченное на заключение Исполнителем_mean'].dt.strftime('%H:%M:%S')
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_top_f


# ---------------------------------------------------------------------
#   Average time of approving
# ---------------------------------------------------------------------
def average_f_apr_r(df_temp_f, date_start, date_end):
    df_top_f = pd.DataFrame()
    try:
        #   Clean
        df_temp_f.dropna()

        #   Changing type of some fields (from Object to Date)
        df_temp_f['Дата заключения УМОД'] = df_temp_f['Дата заключения УМОД'].astype("datetime64[ns]")

        #   Reduction the range of date
        df_top_f = df_temp_f[(df_temp_f['Дата заключения УМОД'] >= date_start + ' 00:00:00') &
                             (df_temp_f['Дата заключения УМОД'] <= date_end + ' 23:59:59')]
        df_top_f = df_top_f[["Уникальность обеспечения", "Филиал исполнитель ДМОД",
                             "Время затраченное на согласование куратором"]]

        #   Formatting to datetime
        df_top_f['Время затраченное на согласование куратором'] = pd.to_datetime(
            df_top_f['Время затраченное на согласование куратором'], format='%H:%M:%S')

        #   Average time
        df_top_f = df_top_f.groupby('Филиал исполнитель ДМОД') \
            .agg({'Время затраченное на согласование куратором': ['mean']}).reset_index()

        #   Formatting to string
        df_top_f.columns = df_top_f.columns.map('_'.join)
        df_top_f['Время затраченное на согласование куратором_mean'] = df_top_f[
            'Время затраченное на согласование куратором_mean'].dt.strftime('%H:%M:%S')
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_top_f


# ---------------------------------------------------------------------
#   Average time of approving, according categories
# ---------------------------------------------------------------------
def average_f_apr_unique(df_temp_f, date_start, date_end):
    df_top_f = pd.DataFrame()
    try:
        #   Clean NaN
        df_temp_f.dropna()

        #   Changing type of some fields (from Object to Date)
        df_temp_f['Дата заключения УМОД'] = df_temp_f['Дата заключения УМОД'].astype("datetime64[ns]")

        #   Reduction the range of date
        df_top_f = df_temp_f[(df_temp_f['Дата заключения УМОД'] >= date_start + ' 00:00:00') &
                             (df_temp_f['Дата заключения УМОД'] <= date_end + ' 23:59:59')]
        df_top_f = df_top_f[["Уникальность обеспечения", "Филиал исполнитель ДМОД",
                             "Время затраченное на согласование куратором"]]

        #   Formatting to datetime
        df_top_f['Время затраченное на согласование куратором'] = pd.to_datetime(
            df_top_f['Время затраченное на согласование куратором'], format='%H:%M:%S')

        #   Average time
        df_top_f = df_top_f.groupby(['Уникальность обеспечения', 'Филиал исполнитель ДМОД']) \
            .agg({'Время затраченное на согласование куратором': ['mean']}).reset_index()

        #   Formatting to string
        df_top_f.columns = df_top_f.columns.map('_'.join)
        df_top_f['Время затраченное на согласование куратором_mean'] = df_top_f[
            'Время затраченное на согласование куратором_mean'].dt.strftime('%H:%M:%S')
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_top_f


# ---------------------------------------------------------------------
#   Return number of planned cases
# ---------------------------------------------------------------------
def calc_plan(df_temp, date_temp):
    #   Init
    number_of_plan = pd.DataFrame()
    try:
        #   Calculate month
        datetime_object = datetime.strptime(date_temp, '%d.%m.%Y')
        date_month = datetime_object.month

        #   Init variable
        date_month_str = ''
        match date_month:
            case 1:
                date_month_str = ['январь']
            case 2:
                date_month_str = ['январь', 'февраль']
            case 3:
                date_month_str = ['январь', 'февраль', 'март']
            case 4:
                date_month_str = ['январь', 'февраль', 'март', 'апрель']
            case 5:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май']
            case 6:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь']
            case 7:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль']
            case 8:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август']
            case 9:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь']
            case 10:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь']
            case 11:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь', 'ноябрь']
            case 12:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь', 'ноябрь', 'декабрь']

        #   Reduce number of columns
        df_temp = df_temp[
            ['Филиал исполнитель ДМОД', 'График (месяц)', 'График скорректированный (месяц)', 'Статус залога 01.2023']]

        #   Calculate number of planning cases
        df_temp = df_temp[((df_temp['График скорректированный (месяц)'].str.lower().isin(date_month_str))
                           | ((pd.isnull(df_temp['График скорректированный (месяц)']))
                              & (df_temp['График (месяц)'].str.lower().isin(date_month_str))))
                          & (df_temp['Статус залога 01.2023'] == 'Принят к учету')]
        number_of_plan = df_temp.groupby('Филиал исполнитель ДМОД').count().reset_index()
        number_of_plan = number_of_plan[['Филиал исполнитель ДМОД', 'График (месяц)']]

        #   Rename result column
        number_of_plan.rename(columns={'График (месяц)': 'План'}, inplace=True)
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return number_of_plan


# ---------------------------------------------------------------------
# Return year monitoring plan
# ---------------------------------------------------------------------
def calc_plan_year(df_temp):
    df_plan_year_temp = pd.DataFrame()
    try:
        #   Reduce number of columns
        df_temp = df_temp[
            ['Филиал исполнитель ДМОД', 'График (месяц)', 'График скорректированный (месяц)', 'Статус залога 01.2023']]

        #   Aggregate data in new column Месяц
        df_temp['Месяц'] = np.where(pd.notna(df_temp['График скорректированный (месяц)']),
                                    df_temp['График скорректированный (месяц)'], df_temp['График (месяц)'])

        #   Calculate total number of monitoring cases
        df_plan_year_temp = df_temp.groupby('Месяц').count().reset_index()

        #   Prepare some Set to arrange names of months
        monthgrades = {'январь': 1, 'февраль': 2, 'март': 3, 'апрель': 4, 'май': 5, 'июнь': 6, 'июль': 7,
                       'август': 8, 'сентябрь': 9, 'октябрь': 10, 'ноябрь': 11, 'декабрь': 12}

        #   Function to sort by month
        def month_sort(series):
            return series.apply(lambda x: monthgrades.get(x))

        #   Sorting
        df_plan_year_temp = df_plan_year_temp.sort_values(by=["Месяц"], key=month_sort)

        #   Reduce number of columns
        df_plan_year_temp = df_plan_year_temp[['Месяц', 'График (месяц)']]

        #   Rename some columns
        df_plan_year_temp.rename(columns={'График (месяц)': 'Количество'}, inplace=True)
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_plan_year_temp


# ---------------------------------------------------------------------
# Return completed monitoring cases
# ---------------------------------------------------------------------
def calc_compl_year(df_temp):
    df_copl_year_temp = pd.DataFrame()
    try:
        #   Reduce number of columns
        df_temp = df_temp[
            ['Филиал исполнитель ДМОД', 'График (месяц)', 'График скорректированный (месяц)', 'Статус залога 01.2023']]

        #   Aggregate data in new column Месяц
        df_temp['Месяц'] = np.where(pd.notna(df_temp['График скорректированный (месяц)']),
                                    df_temp['График скорректированный (месяц)'], df_temp['График (месяц)'])

        #   Calculate total number of monitoring cases
        df_copl_year_temp = df_temp.groupby('Месяц').count().reset_index()

        #   Prepare some Set to arrange names of months
        monthgrades = {'январь': 1, 'февраль': 2, 'март': 3, 'апрель': 4, 'май': 5, 'июнь': 6, 'июль': 7,
                       'август': 8, 'сентябрь': 9, 'октябрь': 10, 'ноябрь': 11, 'декабрь': 12}

        #   Sort by month
        def month_sort(series):
            return series.apply(lambda x: monthgrades.get(x))

        df_copl_year_temp = df_copl_year_temp.sort_values(by=["Месяц"], key=month_sort)

        #   Reduce number of columns
        df_copl_year_temp = df_copl_year_temp[['Месяц', 'График (месяц)']]

        #   Rename result column
        df_copl_year_temp.rename(columns={'График (месяц)': 'Количество'}, inplace=True)
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return df_copl_year_temp


# ---------------------------------------------------------------------
# Return number of expired
# ---------------------------------------------------------------------
def calc_expired(df_plan_temp, df_calc_temp, month_temp):
    number_exp = 0
    try:
        #   Concat two dataframes by Mounth
        df_expired = pd.merge(df_plan_temp, df_calc_temp, how='left', left_on='Месяц',
                              right_on='Месяц')

        #   Fill blank cells by zero
        df_expired['Количество_y'] = df_expired['Количество_y'].fillna(0)

        #   Subtraction two columns to calculate expired cases
        df_expired['Просрочено'] = df_expired['Количество_x'] - df_expired['Количество_y']

        date_month_str = []
        match month_temp:
            case '01':
                date_month_str = []
            case '02':
                date_month_str = ['январь']
            case '03':
                date_month_str = ['январь', 'февраль']
            case '04':
                date_month_str = ['январь', 'февраль', 'март']
            case '05':
                date_month_str = ['январь', 'февраль', 'март', 'апрель']
            case '06':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май']
            case '07':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь']
            case '08':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль']
            case '09':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август']
            case '10':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь']
            case '11':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь']
            case '12':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь', 'ноябрь']

        #   Filter
        df_expired = df_expired[df_expired['Месяц'].str.lower().isin(date_month_str)]

        #   Sum expired
        number_exp = df_expired['Просрочено'].sum()
        number_exp = int(number_exp)
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")

    return number_exp


# ---------------------------------------------------------------------
# Return list of expired cases
# ---------------------------------------------------------------------
def calc_list_expired(df_plan_temp, month_temp):
    #   Reduce number of columns
    list_expired = df_plan_temp[['Код объекта залога', 'Класс обеспечения', 'Наименование заемщика Сцепка',
                                 'Филиал исполнитель ДМОД', 'График скорректированный (месяц)', 'Статус ДМОД']]
    date_month_str = []
    try:
        match month_temp:
            case '01':
                date_month_str = []
            case '02':
                date_month_str = ['январь']
            case '03':
                date_month_str = ['январь', 'февраль']
            case '04':
                date_month_str = ['январь', 'февраль', 'март']
            case '05':
                date_month_str = ['январь', 'февраль', 'март', 'апрель']
            case '06':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май']
            case '07':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь']
            case '08':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль']
            case '09':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август']
            case '10':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь']
            case '11':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь']
            case '12':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь', 'ноябрь']

        #   Filter
        list_expired = list_expired[list_expired['График скорректированный (месяц)'].str.lower().isin(date_month_str)]

        #   Final result list
        list_expired = list_expired[pd.isnull(list_expired['Статус ДМОД'])]
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return list_expired


# ---------------------------------------------------------------------
# Return top expired cases
# ---------------------------------------------------------------------
def calc_top_expired(df_plan_temp, month_temp):
    #   Reduce number of columns
    top_expired = df_plan_temp[['Филиал исполнитель ДМОД', 'График скорректированный (месяц)', 'Статус ДМОД']]

    date_month_str = []
    try:
        match month_temp:
            case '01':
                date_month_str = []
            case '02':
                date_month_str = ['январь']
            case '03':
                date_month_str = ['январь', 'февраль']
            case '04':
                date_month_str = ['январь', 'февраль', 'март']
            case '05':
                date_month_str = ['январь', 'февраль', 'март', 'апрель']
            case '06':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май']
            case '07':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь']
            case '08':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль']
            case '09':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август']
            case '10':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь']
            case '11':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь']
            case '12':
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь', 'ноябрь']

        #   Filter
        top_expired = top_expired[top_expired['График скорректированный (месяц)'].str.lower().isin(date_month_str)]

        #   Pick only not completed cases
        top_expired = top_expired[pd.isnull(top_expired['Статус ДМОД'])]

        #   Count number of expired cases
        top_expired = top_expired.groupby('Филиал исполнитель ДМОД').count().reset_index()

        #   Reduce number of columns
        top_expired = top_expired[['Филиал исполнитель ДМОД', 'График скорректированный (месяц)']]

        #   Rename column
        top_expired.rename(columns={'График скорректированный (месяц)': 'Просрочено'}, inplace=True)
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return top_expired


# ---------------------------------------------------------------------
# Return number of completed cases
# ---------------------------------------------------------------------
def calc_completed(df_temp, date_temp):
    number_compl = pd.DataFrame()
    try:
        #   Create datetime object from string
        datetime_object = datetime.strptime(date_temp, '%d.%m.%Y')

        #   Choose only month
        date_month = datetime_object.month

        date_month_str = []
        match date_month:
            case 1:
                date_month_str = ['январь']
            case 2:
                date_month_str = ['январь', 'февраль']
            case 3:
                date_month_str = ['январь', 'февраль', 'март']
            case 4:
                date_month_str = ['январь', 'февраль', 'март', 'апрель']
            case 5:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май']
            case 6:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь']
            case 7:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль']
            case 8:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август']
            case 9:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь']
            case 10:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь']
            case 11:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь', 'ноябрь']
            case 12:
                date_month_str = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь',
                                  'октябрь', 'ноябрь', 'декабрь']

        #   Reduce number of columns
        df_temp = df_temp[
            ['Филиал исполнитель ДМОД', 'График (месяц)', 'График скорректированный (месяц)', 'Статус ДМОД',
             'Статус залога 01.2023']]

        #   Choose only completed cases
        df_temp = df_temp[((df_temp['График скорректированный (месяц)'].str.lower().isin(date_month_str))
                           | ((pd.isnull(df_temp['График скорректированный (месяц)']))
                              & (df_temp['График (месяц)'].str.lower().isin(date_month_str))))]

        #   Count completed cases by branches
        number_compl = df_temp.groupby('Филиал исполнитель ДМОД').count().reset_index()

        #   Reduce number of columns
        number_compl = number_compl[['Филиал исполнитель ДМОД', 'График (месяц)']]

        #   Rename result column
        number_compl.rename(columns={'График (месяц)': 'Выполнено'}, inplace=True)
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле График планового мониторинга 2023.xlsx, "
              "либо в файле Новый журнал по заявкам 2023.xlsm")
    return number_compl


# ---------------------------------------------------------------------
#    Process file "Новый журнал по заявкам.xlsm" to make list of implementers and approvers
# ---------------------------------------------------------------------
def journal_request(df1, date_start, date_end):
    df_merge_ei = pd.DataFrame()
    try:
        #   Changing type of some fields (from Object to Date)
        df1['Дата заключения УМОД'] = df1['Дата заключения УМОД'].astype("datetime64[ns]")

        #   Reduction the range of date
        df1 = df1[(df1['Дата заключения УМОД'] >= date_start + ' 00:00:00') &
                  (df1['Дата заключения УМОД'] <= date_end + ' 23:59:59')]

        #   Calculate number of implementers
        df_impl1 = df1[["Исполнитель", "Дата заключения УМОД"]].groupby(['Исполнитель']).count()

        #   Rename some columns
        df_impl1.rename(columns={'Дата заключения УМОД': 'Исполнитель'}, inplace=True)

        #   Calculate number of approvers
        df_approver1 = df1[["Согласующий Куратор", "Дата заключения УМОД"]].groupby('Согласующий Куратор').count()

        #   Rename some columns
        df_approver1.rename(columns={'Дата заключения УМОД': 'Согласующий'}, inplace=True)

        #   Merge two DataFrames
        df_merge1 = pd.concat([df_impl1, df_approver1], axis=1, join="outer")
        df_merge_s = df_merge1.assign(DateStart=date_start)
        df_merge_ei = df_merge_s.assign(DateEnd=date_end)
    except Exception as e:
        print(e)
        print("Неправильный формат в ячейке в файле Новый журнал по заявкам 2023.xlsm")
    return df_merge_ei


# ---------------------------------------------------------------------
#    Writing temp table in to CSV file "Мониторинг 2023.csv"
# ---------------------------------------------------------------------
def write_to_csv(df_work):
    try:
        #   Writing DataFrames into csv file (you can change directory and name of file)
        df_work.to_csv('\\\BCCFS-HQ\\reserve_data\отчеты УКР ДКР\ГО\Заявки на согласование ОМОД\Мониторинг 2023.csv',
                       sep=';', encoding='utf-16', index_label='Ф.И.О.')
        # df_work.to_csv('Мониторинг 2023.csv', sep=';', encoding='utf-16', index_label='Ф.И.О.')
        print("File 'Мониторинг 2023.csv' successfully created")
    except Exception as e:
        print(e)
        print("Can not write to a file 'Мониторинг 2023.csv'")


# ---------------------------------------------------------------------
#    Writing some stuff in to file "notes.txt"
# ---------------------------------------------------------------------
def write_to_note(text):
    try:
        #   Choose encoding="ISO-8859-5" for Linux, and utf-16 for Windows
        #   For Linux don't forget uncomment line "value = value.encode()"
        with open("note.txt", mode="w", encoding="utf-16") as file:
            if text:
                file.write(text)
    except Exception as e:
        print(e)
        print("Can't write to a file 'note.txt'")


# ---------------------------------------------------------------------
#    Reading from file "notes.txt"
# ---------------------------------------------------------------------
def read_from_note():
    content = ''
    try:
        #   Choose encoding="ISO-8859-5" for Linux, and utf-16 for Windows
        #   For Linux don't forget uncomment line "value = value.encode()"
        with open("note.txt", mode="r", encoding="utf-16") as file:
            content = file.read()
    except Exception as e:
        print(e)
        print("Can't read file 'note.txt'")
    return content


# -------------------------------------------------------------------------
#   Calculate, query ...
# -------------------------------------------------------------------------
def make_calc(filter_temp, filter_branch):
    global number_plan, top_plan_branch, number_completed, top_completed_branch, top_remained_branch, \
        number_expired, top_expired_branch, list_expired_branch, df_top_filials, \
        number_masssegment_completed, number_unique_completed, number_interval_completed, \
        df_top_leaders, df_top_branches, plan_on_date, df_average_filials_r, df_average_filials_unique, \
        df_plan_year, df_compl_year, df_plan_completed, df_plan_year_branch, df_compl_year_branch, \
        df_average_f_apr_r, df_average_f_apr_unique, completed_on_date, df_list_branch, date_string_today

    # -------------------------------------------------------------------------
    #   Load files
    # -------------------------------------------------------------------------
    # filename1 = "\\\BCCFS-HQ\\reserve_data\отчеты УКР ДКР\ГО\Заявки на согласование ОМОД\Новый журнал по заявкам 2023.xlsm"
    filename1 = "\\\\10.15.129.60\\reserve_data\отчеты УКР ДКР\ГО\Заявки на согласование ОМОД\Новый журнал по заявкам 2023.xlsm"
    # filename1 = "/mnt/share1/Новый журнал по заявкам 2023.xlsm"
    # filename1 = "Новый журнал по заявкам 2023.xlsm"
    df_new_jornal = load_new_jornal(filename1)

    # filename3 = "\\\BCCFS-HQ\\reserve_data\отчеты УКР ДКР\Филиалы\\1. Мониторинг\\4. Мониторинг 2023\График планового мониторинга 2023.xlsx"
    filename3 = "\\\\10.15.129.60\\reserve_data\отчеты УКР ДКР\Филиалы\\1. Мониторинг\\4. Мониторинг 2023\График планового мониторинга 2023.xlsx"
    # filename3 = "/mnt/share2/График планового мониторинга 2023.xlsx"
    # filename3 = "График планового мониторинга 2023.xlsx"
    df_mon = load_mon(filename3)

    #   Initializing range
    date_rep = date.today()
    year_rep = date_rep.strftime("%Y")
    date_string_start = "01.01." + year_rep
    date_string_end = "31.12." + year_rep
    date_string_today = 'Экран выполнения мониторинга объектов залога в ' + year_rep + ' г. на : ' \
                        + date_rep.strftime('%d.%m.%Y')

    # ---------------------------------------------------------------------
    #   Make main dataframe
    # ---------------------------------------------------------------------
    #   Merge two dataframes by column="Код объекта залога"
    df_main_jornal = pd.merge(df_mon, df_new_jornal, how='left', left_on='Код объекта залога', right_on=' ID залога')

    #   Pick columns
    df_main_jornal = df_main_jornal[['Код объекта залога', 'Наименование заемщика Сцепка',
                                     'Класс обеспечения', 'Наименование вида обеспечения', 'Месторасположение (Адрес)',
                                     'Стоимость НОК либо Банка', 'Филиал исполнитель ДМОД', 'Статус залога 01.2023',
                                     'График (месяц)', 'График скорректированный (месяц)', 'Статус ДМОД',
                                     'Ф.И.О. исполнителя', 'Ф.И.О. согласующего',
                                     'Дата заключения ДМОД ', 'Дата заключения УМОД',
                                     'Цель составления заключения', 'Филиал Банка (рассмотрения заявки)',
                                     'Исполнитель', 'Согласующий Куратор', 'Согласующий Начальник отдела',
                                     'Согласующий Начальник Управления', 'Уникальность обеспечения',
                                     'Время затраченное на заключение Исполнителем',
                                     'Время затраченное на согласование куратором',
                                     'В срок/просрочка по нормативному времени исполнителя',
                                     'В срок/просрочка по нормативному времени согласующего куратора']]

    df_main_jornal = df_main_jornal[(df_main_jornal['Статус залога 01.2023'].str.contains('Принят'))]

    #   Calc main plan in to date
    main_plan_on_date = calc_plan(df_main_jornal, date_string_end)

    number_plan = main_plan_on_date['План'].sum()
    top_plan_branch = main_plan_on_date.sort_values('План', ascending=False).head(3)
    if filter_temp == "almaty":
        number_plan_d = main_plan_on_date[
            main_plan_on_date['Филиал исполнитель ДМОД'].str.contains('Алматы')]
        number_plan = number_plan_d['План'].sum()
        top_plan_branch = number_plan_d.sort_values('План', ascending=False).head(3)
    elif filter_temp == "branch":
        number_plan_d = main_plan_on_date[
            ~main_plan_on_date['Филиал исполнитель ДМОД'].str.contains('Алматы')]
        number_plan = number_plan_d['План'].sum()
        top_plan_branch = number_plan_d.sort_values('План', ascending=False).head(3)

    #   Calc completed cases
    df_temp_completed = df_main_jornal[
        (df_main_jornal['Дата заключения УМОД'].notna()) | (df_main_jornal['Статус ДМОД'].notna())]
    df_temp_completed = df_temp_completed[
        (df_temp_completed['Статус ДМОД'].str.contains('исполне'))
        | (df_temp_completed['Статус ДМОД'].str.contains('не требует'))
        | (df_temp_completed['Статус ДМОД'].str.contains('высвобождены'))
        | (df_temp_completed['Статус ДМОД'].str.contains('Заключ'))
        | (df_temp_completed['Статус ДМОД'].str.contains('согласован'))
        | (df_temp_completed['Дата заключения УМОД'].notna())]
    df_temp_completed = df_temp_completed[df_temp_completed['Статус ДМОД'] != 'залог в г.Тараз']
    main_completed_on_date = calc_completed(df_temp_completed, date_string_end)

    #   Calc remained cases
    df_substr_ = pd.merge(main_plan_on_date, main_completed_on_date, how='inner', on='Филиал исполнитель ДМОД')
    df_substr_['Осталось'] = df_substr_['План'] - df_substr_['Выполнено']
    df_substr_ = df_substr_[['Филиал исполнитель ДМОД', 'Осталось']]
    top_remained_branch = df_substr_.sort_values('Осталось', ascending=False).head(3)

    #   Calc list of expired cases
    list_expired_branch = calc_list_expired(df_main_jornal, date_rep.strftime("%m"))

    #   Calc top of expired cases
    top_expired_branch = calc_top_expired(df_main_jornal, date_rep.strftime("%m"))
    top_expired_branch = top_expired_branch.sort_values('Просрочено', ascending=False).head(3)

    number_completed = main_completed_on_date['Выполнено'].sum()
    top_completed_branch = main_completed_on_date.sort_values('Выполнено', ascending=False).head(3)
    if filter_temp == "almaty":
        number_completed_d = main_completed_on_date[
            main_completed_on_date['Филиал исполнитель ДМОД'].str.contains('Алматы')]
        number_completed = number_completed_d['Выполнено'].sum()
        top_completed_branch = number_completed_d.sort_values('Выполнено', ascending=False).head(3)

        df_substr_ = df_substr_[df_substr_['Филиал исполнитель ДМОД'].str.contains('Алматы')]
        top_remained_branch = df_substr_.sort_values('Осталось', ascending=False).head(3)

        top_expired_branch = top_expired_branch[top_expired_branch['Филиал исполнитель ДМОД'].str.contains('Алматы')]
        top_expired_branch = top_expired_branch.sort_values('Просрочено', ascending=False).head(3)
    elif filter_temp == "branch":
        number_completed_d = main_completed_on_date[
            ~main_completed_on_date['Филиал исполнитель ДМОД'].str.contains('Алматы')]
        number_completed = number_completed_d['Выполнено'].sum()
        top_completed_branch = number_completed_d.sort_values('Выполнено', ascending=False).head(3)
        df_substr_ = df_substr_[~df_substr_['Филиал исполнитель ДМОД'].str.contains('Алматы')]
        top_remained_branch = df_substr_.sort_values('Осталось', ascending=False).head(3)
        top_expired_branch = top_expired_branch[~top_expired_branch['Филиал исполнитель ДМОД'].str.contains('Алматы')]
        top_expired_branch = top_expired_branch.sort_values('Просрочено', ascending=False).head(3)

    #   Write to file (work-spreadsheet)
    write_to_csv(df_main_jornal)

    #   Filter by types of query (branches, Almaty etc)
    if filter_temp == "almaty":
        df_main_jornal = df_main_jornal[df_main_jornal['Филиал исполнитель ДМОД'].str.contains('Алматы')]
    elif filter_temp == "branch":
        df_main_jornal = df_main_jornal[~df_main_jornal['Филиал исполнитель ДМОД'].str.contains('Алматы')]

    #   Make list of "completed"
    df_main_jornal_completed = df_main_jornal[
        (df_main_jornal['Дата заключения УМОД'].notna()) | (df_main_jornal['Статус ДМОД'].notna())]
    df_main_jornal_completed = df_main_jornal_completed[
        (df_main_jornal_completed['Статус ДМОД'].str.contains('исполне'))
        | (df_main_jornal_completed['Статус ДМОД'].str.contains('не требует'))
        | (df_temp_completed['Статус ДМОД'].str.contains('высвобождены'))
        | (df_temp_completed['Статус ДМОД'].str.contains('Заключ'))
        | (df_temp_completed['Статус ДМОД'].str.contains('согласован'))
        | (df_main_jornal_completed['Дата заключения УМОД'].notna())]
    df_main_jornal_completed = df_main_jornal_completed[df_main_jornal_completed['Статус ДМОД'] != 'залог в г.Тараз']

    # ---------------------------------------------------------------------
    #   Make work dataframes
    # ---------------------------------------------------------------------
    #   Make year total plan dataframe
    df_plan_year = calc_plan_year(df_main_jornal)

    #   Make dataframe of completed monitoring cases
    df_compl_year = calc_compl_year(df_main_jornal_completed)

    #   Calc expired
    number_expired = calc_expired(df_plan_year, df_compl_year, date_rep.strftime("%m"))

    #   Plan and completed by branches
    plan_on_date = calc_plan(df_main_jornal, date_string_end)
    completed_on_date = calc_completed(df_main_jornal_completed, date_string_end)

    #   Make one dataframe
    df_plan_completed = pd.merge(plan_on_date, completed_on_date, how='left', on='Филиал исполнитель ДМОД')

    #   Change NaN to 0.0
    df_plan_completed['Выполнено'] = df_plan_completed['Выполнено'].fillna(0)
    df_plan_completed = df_plan_completed.sort_values(by=['План'], ascending=False)

    #   Calc dataframe - year plan and completed cases by branches
    if filter_branch == "all":
        df_plan_year_branch = calc_plan_year(df_main_jornal)
        df_compl_year_branch = calc_compl_year(df_main_jornal_completed)
    else:
        df_main_jornal_branch = df_main_jornal[df_main_jornal['Филиал исполнитель ДМОД'].str.contains(filter_branch)]
        df_plan_year_branch = calc_plan_year(df_main_jornal_branch)
        df_main_jornal_completed_branch = df_main_jornal_completed[
            df_main_jornal_completed['Филиал исполнитель ДМОД'].str.contains(filter_branch)]
        df_compl_year_branch = calc_compl_year(df_main_jornal_completed_branch)

    #   Make branches list for dropdown element
    df_plan_completed_temp = df_plan_completed
    if filter_temp == "almaty":
        df_plan_completed_temp = df_plan_completed[df_plan_completed['Филиал исполнитель ДМОД'].str.contains('Алматы')]
    elif filter_temp == "branch":
        df_plan_completed_temp = df_plan_completed[~df_plan_completed['Филиал исполнитель ДМОД'].str.contains('Алматы')]
    df_list_branch = df_plan_completed_temp["Филиал исполнитель ДМОД"].sort_values()

    #   Find the leaders
    df_top_leaders = top_leaders(df_main_jornal_completed, date_string_start, date_string_end)
    df_top_leaders = df_top_leaders.sort_values(by=['Количество'], ascending=[False])
    df_top_leaders = df_top_leaders.head(20)

    df_top_branches = top_branches(df_main_jornal_completed, date_string_start, date_string_end)
    df_top_branches = df_top_branches.sort_values(by=['Количество'], ascending=[False])
    df_top_branches = df_top_branches.head(7)

    #   Leaders by branches
    df_top_filials = completed_on_date.sort_values(by=['Выполнено', 'Филиал исполнитель ДМОД'],
                                                   ascending=[False, True]).head(5)

    #   Calculate completed cases by categories
    number_masssegment_completed = len(df_main_jornal[df_main_jornal['Уникальность обеспечения'] == 'Mass segment'])
    number_unique_completed = len(df_main_jornal[df_main_jornal['Уникальность обеспечения'] == 'Unique'])
    number_interval_completed = len(df_main_jornal[df_main_jornal['Уникальность обеспечения'] == 'Interval'])

    # Average time of monitoring
    df_average_filials_r = average_filials_r(df_main_jornal_completed, date_string_start, date_string_end)
    df_average_filials_r = df_average_filials_r.sort_values(by='Время затраченное на заключение Исполнителем')

    # Average time of monitoring, according categories
    df_average_filials_unique = average_filials_unique(df_main_jornal_completed, date_string_start, date_string_end)
    df_average_filials_unique = df_average_filials_unique.sort_values(
        by='Время затраченное на заключение Исполнителем_mean')

    # Average time of approving
    df_average_f_apr_r = average_f_apr_r(df_main_jornal_completed, date_string_start, date_string_end)
    df_average_f_apr_r = df_average_f_apr_r.sort_values(by='Время затраченное на согласование куратором_mean')

    # Average time of approving, according categories
    df_average_f_apr_unique = average_f_apr_unique(df_main_jornal_completed, date_string_start, date_string_end)
    df_average_f_apr_unique = df_average_f_apr_unique.sort_values(by='Время затраченное на согласование куратором_mean')

    # Find the anti-leaders (outsiders) (maybe useful for future)
    df_antitop_filials_r = antitop_filials_r(df_main_jornal_completed, date_string_start, date_string_end)


# ---------------------------------------------------------------------
#   Prepare graphs, bars and figures
# ---------------------------------------------------------------------
def prepare_fig(colors_temp):
    global fig_main_plan_completed, fig_mass_unique_interval, fig_top_filials, fig_plan_compl_year, \
        fig_plan_completed, fig_plan_completed_branch, fig_mass_unique_interval

    #  Pie: Plan and Completed
    fig_main_plan_completed = px.pie(values=[number_plan - number_completed, number_completed],
                                     names=['Не исполнено', 'Исполнено'],
                                     title='План и всего исполнено',
                                     color_discrete_sequence=['rgb(117, 122, 130)', '#8ecf90']
                                     )
    fig_main_plan_completed.update_traces(hole=.65, hoverinfo='label+percent', textinfo='value')
    fig_main_plan_completed.update_layout(
        annotations=[dict(text="Всего " + str(number_plan), showarrow=False,
                          font={'color': colors_temp['text']})],
        plot_bgcolor='rgba(0, 0, 0, 0)', title_x=0.5,
        margin=dict(t=40, b=40, l=40, r=40),
        legend=dict(yanchor="top", y=0.99, xanchor="center", x=0.01),
        paper_bgcolor='rgba(0, 0, 0, 0)',
        font={'color': colors_temp['text']}
    )

    #  Mass, unique, interval
    fig_mass_unique_interval = px.bar(x=["Mass Segment", "Unique", "Interval"],
                                      y=[number_masssegment_completed, number_unique_completed,
                                         number_interval_completed], text_auto='.2s',
                                      title='Общее количество <BR> исполненых по категориям',
                                      color_discrete_sequence=px.colors.sequential.Blugrn)
    fig_mass_unique_interval.update_yaxes(title='Исполнено', visible=True, showticklabels=False)
    fig_mass_unique_interval.update_xaxes(title='Категории', visible=True, showticklabels=True)
    fig_mass_unique_interval.update_layout(plot_bgcolor='rgba(0, 0, 0, 0)', title_x=0.5,
                                           margin=dict(t=50, b=0, l=20, r=20),
                                           paper_bgcolor='rgba(0, 0, 0, 0)',
                                           font={'color': colors_temp['text']})

    #  Top of branches
    fig_top_filials = px.bar(df_top_filials, x="Филиал исполнитель ДМОД", y="Выполнено", text_auto='.2s',
                             height=400, width=400, color_discrete_map={'Исполнено': '#464646'})
    fig_top_filials.update_layout(plot_bgcolor='rgba(0, 0, 0, 0)', paper_bgcolor='rgba(0, 0, 0, 0)',
                                  title='Филиалы лидеры (кол-во исполненных)', title_x=0.5,
                                  font={'color': colors_temp['text']})

    #   Total number of completed monitoring cases
    fig_plan_compl_year = go.Figure()
    fig_plan_compl_year.add_trace(go.Bar(x=df_plan_year["Месяц"], y=df_plan_year["Количество"], name='План',
                                         text=df_plan_year["Количество"]))
    fig_plan_compl_year.add_trace(go.Bar(x=df_compl_year["Месяц"], y=df_compl_year["Количество"], name="Исполнено",
                                         text=df_compl_year["Количество"]))
    fig_plan_compl_year.update_traces(textfont_size=12, textangle=0, textposition="outside",
                                      cliponaxis=False)
    fig_plan_compl_year.update_layout(barmode='group', xaxis_tickangle=-45, height=500, width=700,
                                      title='План мониторинга на год', title_x=0.5,
                                      plot_bgcolor='rgba(0, 0, 0, 0)', paper_bgcolor='rgba(0, 0, 0, 0)',
                                      font={'color': colors_temp['text']})

    #   Completed monitoring cases by branches
    fig_plan_completed = go.Figure()
    fig_plan_completed.add_trace(go.Bar(x=df_plan_completed["Филиал исполнитель ДМОД"], y=df_plan_completed["План"],
                                        name="План"))
    fig_plan_completed.add_trace(
        go.Bar(x=df_plan_completed["Филиал исполнитель ДМОД"], y=df_plan_completed["Выполнено"],
               name="Исполнено", text=df_plan_completed["Выполнено"]))
    fig_plan_completed.update_traces(textfont_size=12, textangle=0, textposition='outside', cliponaxis=False)
    fig_plan_completed.update_layout(barmode='overlay', xaxis_tickangle=-45, height=500, width=700,
                                     title='Исполнено объектов к плану по всему году', title_x=0.5,
                                     plot_bgcolor='rgba(0, 0, 0, 0)', paper_bgcolor='rgba(0, 0, 0, 0)',
                                     font={'color': colors_temp['text']})

    #   Completed cases by branch
    fig_plan_completed_branch = go.Figure()
    fig_plan_completed_branch.add_trace(
        go.Bar(x=df_plan_year_branch["Месяц"], y=df_plan_year_branch["Количество"], name='План',
               text=df_plan_year_branch["Количество"]))
    fig_plan_completed_branch.add_trace(
        go.Bar(x=df_compl_year_branch["Месяц"], y=df_compl_year_branch["Количество"], name="Исполнено",
               text=df_compl_year_branch["Количество"]))
    fig_plan_completed_branch.update_traces(textfont_size=12, textangle=0, textposition="outside",
                                            cliponaxis=False)
    fig_plan_completed_branch.update_layout(barmode='group', xaxis_tickangle=-45, height=500, width=700,
                                            title='План мониторинга на год по выбранному филиалу', title_x=0.5,
                                            plot_bgcolor='rgba(0, 0, 0, 0)', paper_bgcolor='rgba(0, 0, 0, 0)',
                                            font={'color': colors_temp['text']})


colors = {
    'background': '#c4e6c5',
    'tabletext': '#2d124d',
    'text': '#2d124d'
}
make_calc("all", "all")
prepare_fig(colors)
app = dash.Dash(__name__, external_stylesheets=[dbc.themes.ZEPHYR])
app.title = "Мониторинг"

#   Title of dashboard
header = html.H2(id='output-container-title', style={'textAlign': 'center'})

app.layout = dbc.Container([
    html.Div(style={
        'background-image': 'url("/assets/background2.jpg")',
    }, children=[
        header,
        dbc.Row([
            #   First column (width 2)
            dbc.Col([
                dbc.Card([
                    dcc.Graph(
                        id='graph-main_plan_completed',
                        figure=fig_main_plan_completed
                    ),
                    dbc.Label("Фильтр:", style={'color': colors['text'], 'margin-left': '15px'}),
                    dcc.Dropdown([
                        {'label': 'Вместе г.Алматы + филиальная сеть', 'value': 'all'},
                        {'label': 'Филиальная сеть', 'value': 'branch'},
                        {'label': 'г. Алматы', 'value': 'almaty'}
                    ],
                        'all', id='filter-dropdown',
                        clearable=False, style={'color': "#111111", 'verticalAlign': 'bottom',
                                                'margin-left': '5px', 'margin-right': '15px'}),
                    html.Br(),
                    html.Button(id='submit-button-state', n_clicks=0, children='Обновить',
                                type="button",
                                style={'font-size': '12px', 'width': '140px', 'height': '35px',
                                       'display': 'inline-block',
                                       'margin-left': '10px'
                                       }
                                ),
                    dcc.Graph(
                        id='graph-mass_unique_interval',
                        figure=fig_mass_unique_interval
                    ),
                    html.Br(),
                    html.Img(src='/assets/logo.jpg'),
                ], color="success", outline=True),
            ], width=2),
            #   Second column (width 10)
            dbc.Col([
                dbc.Card([
                    dbc.Row([
                        dbc.Col(
                            dbc.Card([
                                dbc.CardHeader("План"),
                                dbc.CardBody(
                                    [
                                        dbc.Row([
                                            dbc.Col([
                                                html.H1(html.B(id='card_plan')),
                                            ]),
                                            dbc.Col([
                                                dash_table.DataTable(style_data={'whiteSpace': 'normal',
                                                                                 'height': 'auto',
                                                                                 # 'border': 'none'
                                                                                 },
                                                                     id='output-data-table-plan',
                                                                     data=top_plan_branch.to_dict('records'),
                                                                     style_header={'fontWeight': 'bold'},
                                                                     style_cell={'textAlign': 'center',
                                                                                 'backgroundColor': 'transparent',
                                                                                 'color': colors['tabletext'],
                                                                                 'font_size': '12px',
                                                                                 'whiteSpace': 'normal',
                                                                                 'height': 'auto'},
                                                                     columns=[{'id': 'Филиал исполнитель ДМОД',
                                                                               'name': ['Филиал']},
                                                                              {'id': 'План', 'name': ['План']}],
                                                                     ),
                                            ])
                                        ])
                                    ]
                                ),
                            ], color="success", outline=False, inverse=True),
                            width=3,
                        ),
                        dbc.Col(
                            dbc.Card([
                                dbc.CardHeader("Исполнено"),
                                dbc.CardBody(
                                    [
                                        dbc.Row([
                                            dbc.Col([
                                                html.H1(html.B(id='card_completed')),
                                            ]),
                                            dbc.Col([
                                                dash_table.DataTable(style_data={'whiteSpace': 'normal',
                                                                                 'height': 'auto',
                                                                                 # 'border': 'none'
                                                                                 },
                                                                     id='output-data-table-comleted',
                                                                     data=top_completed_branch.to_dict('records'),
                                                                     style_header={'fontWeight': 'bold'},
                                                                     style_cell={'textAlign': 'center',
                                                                                 'backgroundColor': 'transparent',
                                                                                 'color': colors['tabletext'],
                                                                                 'font_size': '12px',
                                                                                 'whiteSpace': 'normal',
                                                                                 'height': 'auto'},
                                                                     columns=[{'id': 'Филиал исполнитель ДМОД',
                                                                               'name': ['Филиал']},
                                                                              {'id': 'Выполнено',
                                                                               'name': ['Исполнено']}],
                                                                     ),
                                            ])
                                        ])
                                    ]
                                ),
                            ], color="info", outline=False, inverse=True),
                            width=3,
                        ),
                        dbc.Col(
                            dbc.Card([
                                dbc.CardHeader("Осталось"),
                                dbc.CardBody(
                                    [
                                        dbc.Row([
                                            dbc.Col([
                                                html.H1(html.B(id='card_left')),
                                            ]),
                                            dbc.Col([
                                                dash_table.DataTable(style_data={'whiteSpace': 'normal',
                                                                                 'height': 'auto',
                                                                                 },
                                                                     id='output-data-table-remined',
                                                                     data=top_remained_branch.to_dict('records'),
                                                                     style_header={'fontWeight': 'bold'},
                                                                     style_cell={'textAlign': 'center',
                                                                                 'backgroundColor': 'transparent',
                                                                                 'color': colors['tabletext'],
                                                                                 'font_size': '12px',
                                                                                 'whiteSpace': 'normal',
                                                                                 'height': 'auto'},
                                                                     columns=[{'id': 'Филиал исполнитель ДМОД',
                                                                               'name': ['Филиал']},
                                                                              {'id': 'Осталось',
                                                                               'name': ['Осталось']}],
                                                                     ),
                                            ])
                                        ]),
                                    ]
                                ),
                            ], color="warning", outline=False, inverse=True),
                            width=3,
                        ),
                        dbc.Col(
                            dbc.Card([
                                dbc.CardHeader("Просрочено"),
                                dbc.CardBody(
                                    [
                                        dbc.Row([
                                            dbc.Col([
                                                html.H1(html.B(id="card_expired")),
                                            ]),
                                            dbc.Col([
                                                dash_table.DataTable(style_data={'whiteSpace': 'normal',
                                                                                 'height': 'auto',
                                                                                 },
                                                                     id='output-data-table-expired',
                                                                     data=top_expired_branch.to_dict('records'),
                                                                     style_header={'fontWeight': 'bold'},
                                                                     style_cell={'textAlign': 'center',
                                                                                 'backgroundColor': 'transparent',
                                                                                 'color': colors['tabletext'],
                                                                                 'font_size': '12px',
                                                                                 'whiteSpace': 'normal',
                                                                                 'height': 'auto'},
                                                                     columns=[{'id': 'Филиал исполнитель ДМОД',
                                                                               'name': ['Филиал']},
                                                                              {'id': 'Просрочено',
                                                                               'name': ['Просрочено']}],
                                                                     ),
                                                dbc.NavLink("Подробнее", id="expired_nav_link", n_clicks=0),
                                            ])
                                        ]),
                                    ]
                                ),
                            ], color="danger", outline=False, inverse=True),
                            width=3,
                        ),
                    ]),
                ]),
                html.Br(),
                dbc.Card([
                    dbc.Tabs(id='tabs', active_tab='year_plan_tab', children=[
                        dbc.Tab(children=[
                            dcc.Graph(
                                id='graph-plan_compl_year',
                                figure=fig_plan_compl_year
                            )
                        ], label="План на год", tab_id="year_plan_tab"),
                        dbc.Tab([
                            dbc.Row([
                                dbc.Col([
                                    dcc.Graph(
                                        id='graph-top_filials',
                                        figure=fig_top_filials
                                    ),
                                ], align="center"),
                                dbc.Col([
                                    dash_table.DataTable(style_data={'whiteSpace': 'normal', 'height': 'auto'},
                                                         id='output-data-table-r12',
                                                         data=df_top_leaders.to_dict('records'),
                                                         style_header={'backgroundColor': colors['background'],
                                                                       'padding': '10px',
                                                                       'fontWeight': 'bold',
                                                                       'color': colors['tabletext']},
                                                         style_cell={'textAlign': 'center',
                                                                     'backgroundColor': colors['background'],
                                                                     'color': colors['tabletext'],
                                                                     'font_size': '14px', 'whiteSpace': 'normal',
                                                                     'height': 'auto'},
                                                         editable=True,
                                                         sort_action="native",
                                                         sort_mode="multi",
                                                         cell_selectable=True,
                                                         merge_duplicate_headers=True,
                                                         columns=[
                                                             {'id': 'Ф.И.О. исполнителя',
                                                              'name': ['Лидеры', 'Исполнитель']},
                                                             {'id': 'Количество', 'name': ['Лидеры', 'Количество']}
                                                         ]),
                                ])
                            ]),
                        ], label="Лидеры", tab_id="top_tab"),
                        dbc.Tab(children=[
                            dbc.Row([
                                dbc.Col([
                                    dash_table.DataTable(style_data={'whiteSpace': 'normal', 'height': 'auto'},
                                                         id='output-average-r',
                                                         data=df_average_filials_r.to_dict('records'),
                                                         style_header={'backgroundColor': colors['background'],
                                                                       'padding': '10px',
                                                                       'fontWeight': 'bold',
                                                                       'color': colors['tabletext']},
                                                         style_cell={'textAlign': 'center',
                                                                     'backgroundColor': colors['background'],
                                                                     'color': colors['tabletext'],
                                                                     'font_size': '14px', 'whiteSpace': 'normal',
                                                                     'height': 'auto'},
                                                         editable=True,
                                                         sort_action="native",
                                                         sort_mode="multi",
                                                         cell_selectable=True,
                                                         merge_duplicate_headers=True,
                                                         columns=[{'id': 'Филиал исполнитель ДМОД',
                                                                   'name': ['Среднее время подготовки заключения.',
                                                                            'Филиал исполнитель ДМОД']},
                                                                  {'id': 'Время затраченное на заключение Исполнителем',
                                                                   'name': ['Среднее время подготовки заключения.',
                                                                            'Время затраченное на заключение Исполнителем']}
                                                                  ])
                                ]),
                                dbc.Col([
                                    dash_table.DataTable(style_data={'whiteSpace': 'normal', 'height': 'auto'},
                                                         id='output-average-unique',
                                                         data=df_average_filials_unique.to_dict('records'),
                                                         style_header={'backgroundColor': colors['background'],
                                                                       'padding': '10px',
                                                                       'fontWeight': 'bold',
                                                                       'color': colors['tabletext']},
                                                         style_cell={'textAlign': 'center',
                                                                     'backgroundColor': colors['background'],
                                                                     'color': colors['tabletext'],
                                                                     'font_size': '14px', 'whiteSpace': 'normal',
                                                                     'height': 'auto'},
                                                         editable=True,
                                                         sort_action="native",
                                                         sort_mode="multi",
                                                         cell_selectable=True,
                                                         merge_duplicate_headers=True,
                                                         columns=[{'id': 'Уникальность обеспечения_',
                                                                   'name': [
                                                                       'Среднее время подготовки заключения по категориям.',
                                                                       'Категория']},
                                                                  {'id': 'Филиал исполнитель ДМОД_',
                                                                   'name': [
                                                                       'Среднее время подготовки заключения по категориям.',
                                                                       'Филиал исполнитель ДМОД']},
                                                                  {
                                                                      'id': 'Время затраченное на заключение Исполнителем_mean',
                                                                      'name': [
                                                                          'Среднее время подготовки заключения по категориям.',
                                                                          'Время затраченное на заключение Исполнителем']}
                                                                  ])
                                ]),
                            ]),
                            dbc.Row([
                                dbc.Col([
                                    dash_table.DataTable(style_data={'whiteSpace': 'normal', 'height': 'auto'},
                                                         id='output-average-f_apr',
                                                         data=df_average_f_apr_r.to_dict('records'),
                                                         style_header={'backgroundColor': colors['background'],
                                                                       'padding': '10px',
                                                                       'fontWeight': 'bold',
                                                                       'color': colors['tabletext']},
                                                         style_cell={'textAlign': 'center',
                                                                     'backgroundColor': colors['background'],
                                                                     'color': colors['tabletext'],
                                                                     'font_size': '14px', 'whiteSpace': 'normal',
                                                                     'height': 'auto'},
                                                         editable=True,
                                                         sort_action="native",
                                                         sort_mode="multi",
                                                         cell_selectable=True,
                                                         merge_duplicate_headers=True,
                                                         columns=[{'id': 'Филиал исполнитель ДМОД_',
                                                                   'name': [
                                                                       'Среднее время подготовки согласования заключения',
                                                                       'Филиал исполнитель ДМОД']},
                                                                  {
                                                                      'id': 'Время затраченное на согласование куратором_mean',
                                                                      'name': [
                                                                          'Среднее время подготовки согласования заключения',
                                                                          'Время затраченное на согласование куратором']}
                                                                  ])
                                ]),
                                dbc.Col([
                                    dash_table.DataTable(style_data={'whiteSpace': 'normal', 'height': 'auto'},
                                                         id='output-average-f_apr_unique',
                                                         data=df_average_f_apr_unique.to_dict('records'),
                                                         style_header={'backgroundColor': colors['background'],
                                                                       'padding': '10px',
                                                                       'fontWeight': 'bold',
                                                                       'color': colors['tabletext']},
                                                         style_cell={'textAlign': 'center',
                                                                     'backgroundColor': colors['background'],
                                                                     'color': colors['tabletext'],
                                                                     'font_size': '14px', 'whiteSpace': 'normal',
                                                                     'height': 'auto'},
                                                         editable=True,
                                                         sort_action="native",
                                                         sort_mode="multi",
                                                         cell_selectable=True,
                                                         merge_duplicate_headers=True,
                                                         columns=[{'id': 'Уникальность обеспечения_',
                                                                   'name': [
                                                                       'Среднее время согласования заключения по категориям',
                                                                       'Уникальность обеспечения']},
                                                                  {'id': 'Филиал исполнитель ДМОД_',
                                                                   'name': [
                                                                       'Среднее время согласования заключения по категориям',
                                                                       'Филиал исполнитель ДМОД']},
                                                                  {
                                                                      'id': 'Время затраченное на согласование куратором_mean',
                                                                      'name': [
                                                                          'Среднее время согласования заключения по категориям',
                                                                          'Время затраченное на согласование куратором']}
                                                                  ])
                                ]),
                            ]),
                        ], label="Среднее время", tab_id="average_time_tab"),
                        dbc.Tab(children=[
                            dbc.Row([
                                dbc.Col([
                                    dcc.Graph(
                                        id='graph-plan-completed',
                                        figure=fig_plan_completed
                                    ),
                                ]),
                                dbc.Col([
                                    dcc.Dropdown(df_list_branch, 'all', id='branch-dropdown',
                                                 clearable=False, style={'color': "#111111", 'verticalAlign': 'bottom',
                                                                         'margin-left': '5px', 'margin-right': '15px'}),
                                    dcc.Graph(
                                        id='graph-plan-completed-branch',
                                        figure=fig_plan_completed_branch
                                    ),
                                ]),
                                dbc.Col([
                                    dash_table.DataTable(style_data={'whiteSpace': 'normal', 'height': 'auto'},
                                                         id='plan_on_date',
                                                         data=plan_on_date.to_dict('records'),
                                                         style_header={'backgroundColor': colors['background'],
                                                                       'padding': '10px',
                                                                       'fontWeight': 'bold',
                                                                       'color': colors['tabletext']},
                                                         style_cell={'textAlign': 'center',
                                                                     'backgroundColor': colors['background'],
                                                                     'color': colors['tabletext'],
                                                                     'font_size': '14px', 'whiteSpace': 'normal',
                                                                     'height': 'auto'},
                                                         editable=True,
                                                         sort_action="native",
                                                         sort_mode="multi",
                                                         cell_selectable=True,
                                                         merge_duplicate_headers=True,
                                                         columns=[{'id': 'Филиал исполнитель ДМОД',
                                                                   'name': ['План мониторинга за год',
                                                                            'Филиал исполнитель ДМОД']},
                                                                  {'id': 'План',
                                                                   'name': ['План мониторинга за год', 'План']}
                                                                  ])
                                ]),
                                dbc.Col([
                                    dash_table.DataTable(style_data={'whiteSpace': 'normal', 'height': 'auto'},
                                                         id='completed_on_date_id',
                                                         data=completed_on_date.to_dict('records'),
                                                         style_header={'backgroundColor': colors['background'],
                                                                       'padding': '10px',
                                                                       'fontWeight': 'bold',
                                                                       'color': colors['tabletext']},
                                                         style_cell={'textAlign': 'center',
                                                                     'backgroundColor': colors['background'],
                                                                     'color': colors['tabletext'],
                                                                     'font_size': '14px', 'whiteSpace': 'normal',
                                                                     'height': 'auto'},
                                                         editable=True,
                                                         sort_action="native",
                                                         sort_mode="multi",
                                                         cell_selectable=True,
                                                         merge_duplicate_headers=True,
                                                         columns=[{'id': 'Филиал исполнитель ДМОД',
                                                                   'name': ['Исполнено объектов мониторинга',
                                                                            'Филиал исполнитель ДМОД']},
                                                                  {'id': 'Выполнено',
                                                                   'name': ['Исполнено объектов мониторинга',
                                                                            'Исполнено']}
                                                                  ])
                                ]),
                            ]),
                        ], label="Исполнено", tab_id="completed_tab"),
                        dbc.Tab(children=[
                            dash_table.DataTable(style_data={'whiteSpace': 'normal', 'height': 'auto'},
                                                 id='list_expired_id',
                                                 data=list_expired_branch.to_dict('records'),
                                                 style_header={'backgroundColor': colors['background'],
                                                               'padding': '10px',
                                                               'fontWeight': 'bold',
                                                               'color': colors['tabletext']},
                                                 style_cell={'textAlign': 'center',
                                                             'backgroundColor': colors['background'],
                                                             'color': colors['tabletext'],
                                                             'font_size': '14px', 'whiteSpace': 'normal',
                                                             'height': 'auto'},
                                                 editable=True,
                                                 sort_action="native",
                                                 sort_mode="multi",
                                                 cell_selectable=True,
                                                 merge_duplicate_headers=True,
                                                 columns=[{'id': 'Филиал исполнитель ДМОД',
                                                           'name': 'Филиал исполнитель ДМОД'},
                                                          {'id': 'Код объекта залога',
                                                           'name': 'Код объекта залога'},
                                                          {'id': 'Класс обеспечения',
                                                           'name': 'Класс обеспечения', },
                                                          {'id': 'Наименование заемщика Сцепка',
                                                           'name': 'Наименование заемщика'},
                                                          {'id': 'График скорректированный (месяц)',
                                                           'name': 'Месяц'}
                                                          ])
                        ], label='Просрочено', tab_id='expired_tab'),
                        dbc.Tab(children=[
                            dcc.Textarea(
                                id='textarea-example',
                                value=read_from_note(),
                                persistence=True, persistence_type='local',
                                style={'width': '100%', 'height': 300, 'background-color': colors['background'],
                                       'color': colors['tabletext'], 'margin-left': '0px', 'margin-right': '5px'},
                            )
                        ], label="Рекомендации", tab_id="textarea_tab"),
                    ]),
                ], color="success", outline=True),
            ], width=10)
        ]),
    ]),
], fluid=True, className="dbc")


@app.callback(
    Output('output-container-title', 'children'),
    Output('graph-main_plan_completed', 'figure'),
    Output('graph-mass_unique_interval', 'figure'),
    Output('card_plan', 'children'),
    Output('output-data-table-plan', 'data'),
    Output('card_completed', 'children'),
    Output('output-data-table-comleted', 'data'),
    Output('card_left', 'children'),
    Output('output-data-table-remined', 'data'),
    Output('card_expired', 'children'),
    Output('output-data-table-expired', 'data'),
    Output('tabs', 'active_tab'),
    Output('graph-plan_compl_year', 'figure'),
    Output('graph-top_filials', 'figure'),
    Output('output-data-table-r12', 'data'),
    Output('output-average-r', 'data'),
    Output('output-average-unique', 'data'),
    Output('output-average-f_apr', 'data'),
    Output('output-average-f_apr_unique', 'data'),
    Output('plan_on_date', 'data'),
    Output('graph-plan-completed', 'figure'),
    Output('branch-dropdown', 'options'),
    Output('graph-plan-completed-branch', 'figure'),
    Output('completed_on_date_id', 'data'),
    Output('list_expired_id', 'data'),
    Input('tabs', 'active_tab'),
    Input('submit-button-state', 'n_clicks'),
    Input('expired_nav_link', 'n_clicks'),
    Input('textarea-example', 'value'),
    Input('filter-dropdown', 'value'),
    Input('branch-dropdown', 'value')
)
def update_output(active_tab_, n_clicks, nav_clicks, value, main_filter, branch_filter):
    #   encode text for Linux
    # value = value.encode()

    #   Write text from Textarea to file note.txt
    write_to_note(format(value))

    #   Main filter of dashboard, according this filter make calculations
    if main_filter == "branch":
        make_calc("branch", branch_filter)
    elif main_filter == "almaty":
        make_calc("almaty", branch_filter)
    else:
        make_calc("all", branch_filter)

    #   Process link to expired cases list
    if "expired_nav_link" == ctx.triggered_id:
        active_tab_ = 'expired_tab'

    #   Make graphs, histograms
    prepare_fig(colors)

    return date_string_today, fig_main_plan_completed, fig_mass_unique_interval, \
           str(number_plan), top_plan_branch.to_dict('records'), \
           str(number_completed), top_completed_branch.to_dict('records'), \
           str(number_plan - number_completed), top_remained_branch.to_dict('records'), str(number_expired), \
           top_expired_branch.to_dict('records'), active_tab_, \
           fig_plan_compl_year, fig_top_filials, \
           df_top_leaders.to_dict('records'), \
           df_average_filials_r.to_dict('records'), \
           df_average_filials_unique.to_dict('records'), \
           df_average_f_apr_r.to_dict('records'), \
           df_average_f_apr_unique.to_dict('records'), plan_on_date.to_dict('records'), \
           fig_plan_completed, df_list_branch, fig_plan_completed_branch, completed_on_date.to_dict('records'), \
           list_expired_branch.to_dict('records')


# --------------------------------------------------------------------------------------------------------------------
#   Main part
# --------------------------------------------------------------------------------------------------------------------
if __name__ == '__main__':
    app.run_server(debug=True, port='8000', host='0.0.0.0')
