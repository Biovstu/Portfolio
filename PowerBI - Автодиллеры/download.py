#!/usr/bin/env python
# coding: utf-8

# In[1]:


import urllib.request
import pandas as pd
from statsmodels.tsa.ar_model import AutoReg


# Сохранаяем исходные данные из Гугл-таблиц в файл xlsx

# In[2]:


print('Начинаем обновление источника')
print('1 из 8 - Загружаем файл по ссылке')
destination = 'C:\\Users\\gehor\\OneDrive\\Документы\\BI\\Итоговая контрольная работа по блоку специализация\\source.xlsx'
url = 'https://docs.google.com/spreadsheets/d/152JyksagijqyscnrFDc6Ez2VjT5MKNXpDOyc4PRlauw/edit#gid=208646510'
url = url.replace('edit#gid=208646510', 'export?format=xlsx')
urllib.request.urlretrieve(url, destination)
print('Файл по ссылке загружен')


# Читаем листы из листа xlsx в таблицы

# In[3]:


print('2 из 8 - Читаем из файла данные')
marketing = pd.read_excel(destination, sheet_name='Маркетинговые данные')
spravka = pd.read_excel(destination, sheet_name='Справочник')
crm = pd.read_excel(destination, sheet_name='Данные из CRM')
print('Данные загружены')


# Собираем уникальные записи ID пользователей<p>
# И собираем касания по Device Category

# In[4]:


print('3 из 8 - Считаем касания')
clients = marketing['Client ID'].unique()
touch = []
conversion_rate = []
qty_touch = []
touchings = pd.DataFrame({})
for key in clients:
    lst = list(marketing[marketing['Client ID'] == key]['Device Category'])
    touch.append('=>'.join(lst))
    qty_touch.append(len(lst))
    conversion_rate.append(sum(list(marketing[marketing['Client ID'] == key]['Goal Conversion Rate'])))
touchings['Client ID'] = clients
touchings['Touch'] = touch
touchings['QTY Touch'] = qty_touch
touchings['Conversion Rate'] = conversion_rate
touchings['Conversion'] = touchings['Conversion Rate'].apply(lambda x: 1 if x > 0 else 0)
print('Касания определены')


# Разделяем поле Domain на 1.1, 1.2, 2, 3

# In[5]:


def dom(field):
    tmp = field.split('-')
    domains = []
    domains.append(tmp[1])
    if len(tmp) > 2:
        domains.append(tmp[2])
    else:
        domains.append('')
    if '.' in tmp[0]:
        domains += tmp[0].split('.')
    else:
        domains.append(tmp[0])
        domains.append('')
    return domains
print('4 из 8 - Раскрываем домен')
marketing['Domain.2'] = marketing['Domain'].apply(lambda x: dom(x)[0])
marketing['Domain.3'] = marketing['Domain'].apply(lambda x: dom(x)[1])
marketing['Domain.1.1'] = marketing['Domain'].apply(lambda x: dom(x)[2])
marketing['Domain.1.2'] = marketing['Domain'].apply(lambda x: dom(x)[3])
# marketing
print('Домен раскрыт')


# Определяем марку по полям Domain 1.1 и 1.2

# In[6]:


print('5 из 8 - Определяем марку и модель')
marketing['Марка'] = marketing['Domain.1.2']
marketing.loc[marketing['Марка'] == '', 'Марка'] = marketing.loc[marketing['Марка'] == '', 'Domain.1.1']
# marketing


# Ищем упоминание любой модели из справочника в поле Goal Completion Location<p>
# Если есть, то это является моделью для данной записи

# In[7]:


models = spravka['Модель']
def model(field):
    tmp = ''
    for i in models:
        if str(i) in field:
            tmp = str(i)
    return tmp
marketing['Модель'] = marketing['Goal Completion Location'].apply(model)
# marketing
print('Марка и модель определены')


# Подготовливаем Прайс-лист

# In[8]:


print('6 из 8 - Добавляем стоимость авто')
price_list = pd.DataFrame({})
price_list['Модель'] = spravka['Модель']
price_list['Цена'] = spravka['Цена']
# price_list


# Добавляем стоимость

# In[9]:


marketing = marketing.merge(price_list, how='left', on='Модель')
marketing.rename(columns={'Цена':'Стоимость'}, inplace=True)
print('Стоимость авто добавлено')


# In[10]:


crm.rename(columns={'Город':'City', 'Приход к диллеру':'Приход к дилеру'}, inplace=True)
crm.loc[(crm['Просчет стоимости модели'].isnull() | (crm['Просчет стоимости модели'] == '-')),['Просчет стоимости модели']] = 0
crm.loc[(crm['Приход к дилеру'].isnull() | (crm['Приход к дилеру'] == '-')),['Приход к дилеру']] = 0
crm.loc[(crm['Продажа'].isnull() | (crm['Продажа'] == '-') | (crm['Продажа'] == 'Нет данных')),['Продажа']] = 0
crm['ID+City'] = crm['Client ID'] + '+' + crm['City']
crm['Продажа'] = crm['Продажа'].astype('int64')
crm_sales = pd.pivot_table(crm, values='Продажа', index='ID+City', aggfunc=sum).reset_index()
marketing['ID+City'] = marketing['Client ID'] + '+' + marketing['City']
marketing = marketing.merge(crm_sales, how='left', on='ID+City')
marketing.loc[marketing['Модель'] == '', 'Продажа'] = 0
marketing.drop(columns='ID+City', inplace = True)


# In[12]:


print('7 из 8 - Расчитаем прогноз')
forecast = marketing[['Date', 'Конверсия']]
forecast = pd.pivot_table(forecast, values='Конверсия', index='Date', aggfunc=sum).reset_index()
forecast['Day'] = forecast['Date'].apply(lambda x: x.split(' ')[0])
forecast['Month'] = forecast['Date'].apply(lambda x: x.split(' ')[1])
forecast['Year'] = forecast['Date'].apply(lambda x: x.split(' ')[2])
forecast['Month'].replace({'января':'1','февраля':'2'}, inplace=True)
forecast['Year'] = forecast['Year'].astype('int64')
forecast['Month'] = forecast['Month'].astype('int64')
forecast['Day'] = forecast['Day'].astype('int64')
forecast['Факт/Прогноз'] = 'Факт'
forecast.sort_values(['Year', 'Month', 'Day'], axis=0, inplace=True, ignore_index=True)
conversion = forecast['Конверсия']
conv_forcast = model = AutoReg(conversion, lags=12).fit().forecast(steps=12)
conv_forcast = pd.DataFrame(conv_forcast)
conv_forcast.rename(columns={0: 'Конверсия'}, inplace=True)
conv_forcast['Year'] = 2020
conv_forcast['Month'] = 2
conv_forcast['Факт/Прогноз'] = 'Прогноз'
conv_forcast['Day'] = conv_forcast.index - 30
conv_forcast['Date'] = conv_forcast['Day'].astype('str').apply(lambda x: x + ' февраля 2020 г.')
forecast = pd.concat([forecast, conv_forcast])
print('Прогноз готов')


# Перезаписываем исходный файл с 4мя обновленными листами

# In[13]:


print('8 из 8 - Записываем данные в источник')
with pd.ExcelWriter(destination) as writer:
    touchings.to_excel(writer, sheet_name='Касания', index=False)
    crm.to_excel(writer, sheet_name='Данные из CRM', index=False)
    spravka.to_excel(writer, sheet_name='Справочник', index=False)
    marketing.to_excel(writer, sheet_name='Маркетинговые данные', index=False)
    forecast.to_excel(writer, sheet_name='Прогноз', index=False)
print('Источник обновлен')

