#!/usr/bin/env python
# coding: utf-8

# In[50]:


from pathlib import Path
import pandas as pd
import pyecharts as pyecharts
from pyecharts import options as opts
from pyecharts.charts import Bar, Calendar, Tab


# In[17]:


pwd


# In[37]:


pd.read_excel("C:\\Users\\Lenovo\\Desktop\\Python learning notes\\My projects\\Newproject_Pyechart dashboard\\Financial_Data.xlsx")


# In[56]:


excel_file = 'C:\\Users\\Lenovo\\Desktop\\Python learning notes\\My projects\\Newproject_Pyechart dashboard\\Financial_Data.xlsx'
engine = 'openpyxl'
sheet_name = 'Orders'
df = pd.read_excel(excel_file,
                  sheet_name = sheet_name,
                  header = 1,
                  usecols = 'A:T')
df.head()


# In[57]:


df.info()


# In[58]:


df['Month'] = df['Order Date'].dt.month
df.head(3)


# In[66]:


grouped_by_months = df.groupby(by=['Month']).sum()[['Sales', 'Profit']]
grouped_by_months.head(3)


# In[82]:


bar_chart_by_month = (
    Bar()
    .add_xaxis(grouped_by_months.index.tolist())
    .add_yaxis('Sales', grouped_by_months['Sales'].round(0).tolist())
    .add_yaxis('Profit', grouped_by_months['Profit'].round(0).tolist())
    .set_series_opts(label_opts=opts.LabelOpts(position='top'))
    .set_global_opts(
        title_opts=opts.TitleOpts(title='Sales & Profit by month', subtitle='in USD')
    )
)


# In[83]:


bar_chart_by_month.render_notebook()


# In[74]:


grouped_by_subcategory = df.groupby(by='Sub-Category', as_index=False).sum().sort_values(by=['Sales'])
grouped_by_subcategory.head(3)


# In[78]:


bar_chart_by_subcategory = (
    Bar()
    .add_xaxis(grouped_by_subcategory['Sub-Category'].tolist())
    .add_yaxis('Sales', grouped_by_months['Sales'].round(0).tolist())
    .add_yaxis('Profit', grouped_by_months['Profit'].round(0).tolist())
    .reversal_axis()
    .set_series_opts(label_opts=opts.LabelOpts(position='right'))
    .set_global_opts(
        title_opts=opts.TitleOpts(title='Sales & Profit by Sub-Category', subtitle='in USD')
    )
)


# In[79]:


bar_chart_by_subcategory.render_notebook()


# In[86]:


sales_by_date = df.groupby(by='Order Date').sum()[['Sales']].round(0)
sales_by_date = sales_by_date.reset_index()
sales_by_date.head(3)


# In[87]:


data = sales_by_date[['Order Date','Sales']].values.tolist()
data


# In[88]:


max_sales = df['Sales'].max()
min_sales = df['Sales'].min()


# In[99]:


sales_calendar = (
    Calendar()
    .add('',data, calendar_opts=opts.CalendarOpts(range_='2021'))
    .set_global_opts(
        title_opts=opts.TitleOpts(title = 'Sales Calendar', subtitle='in USD'),
        legend_opts=opts.LegendOpts(is_show=False),
        visualmap_opts=opts.VisualMapOpts(
                max_=max_sales,
                min_=min_sales,
                orient='horizontal',
                is_piecewise=False,
                pos_top='230px',
                pos_left='100px',
        )
    )
)


# In[100]:


sales_calendar.render_notebook()


# In[107]:


tab = Tab(page_title='Sales & Profit Overview')
tab.add(bar_chart_by_month, 'Sales & Profit by months')
tab.add(bar_chart_by_subcategory,'Sales & Profit by Sub-Category')
tab.add(sales_calendar, 'Sales Calendar')
tab.render(Path('C:\\Users\\Lenovo\\Desktop\\Python learning notes\\My projects\\Newproject_Pyechart dashboard') / 'Sales & Profit Overview.html')


# In[ ]:




