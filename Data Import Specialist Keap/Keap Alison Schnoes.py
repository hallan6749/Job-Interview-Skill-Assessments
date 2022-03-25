#!/usr/bin/env python
# coding: utf-8

# # Split the Name column in file 1 into First Name and Last Name

# ## Load Libraries

# In[1]:


import pandas as pd


# ## Read csv from excel document file 1 and show first 5 columns

# In[2]:


file1 = pd.read_csv('../Keap/File 1.csv')
file1.head(5)


# ## New data frame with split value columns

# In[3]:


file1new = file1["Name"].str.split(" ", n = 1, expand = True)


# ## Making separate first name column from new data frame

# In[4]:


file1["First Name"] = file1new[0]


# ## Making separate last name column from new data frame

# In[5]:


file1["Last Name"] = file1new[1]


# ## Dropping the old Name column

# In[6]:


file1.drop(columns =["Name"], inplace = True)


# ## Show first 5 columns for first and last name split from file 1

# In[7]:


file1.head()


# # Split the State column in file 1 into City and State

# In[8]:


file1[['City', 'State']] = file1['State'].str.split(',', n=1, expand=True)


# ## Show first 5 columns with the split of city and state

# In[9]:


file1.head()


# # Set the Country column in file 1 as 'United States' for all rows

# ## Find out different values on the Country column to set to United States

# In[10]:


file1.value_counts('Country')


# ## Recode the 28 counts of USA and the 13 counts of US into United States

# In[11]:


def countryFUNCTION(series):
    if series == 'USA':
        return 'United States'
    if series == 'US':
        return 'United States'
    else:
        return series

file1['Country'] = file1['Country'].apply(countryFUNCTION)
        
file1['Country'].value_counts()


# ## Show first 5 columns with the Country renamed as the United States for all rows

# In[12]:


file1.head()


# # Set the Date Created column in file *2* to show as DD/MM/YYYY

# ## Read csv from excel document file 2 and show first 5 columns

# In[13]:


file2 = pd.read_csv('../Keap/File 2.csv')
file2.head(5)


# ## Load Library

# In[14]:


from datetime import datetime


# ## Define a series and then use date formating to change File 2

# In[15]:


def conv_dates_series(df, col, old_date_format, new_date_format):

    file2['Date Created'] = pd.to_datetime(file2['Date Created'], format=old_date_format).dt.strftime(new_date_format)
    
    return file2


# In[16]:


test_df = pd.DataFrame({"file2": ["01-01-0000", "01-01-9999"]})

old_date_format='%m/%d/%Y'
new_date_format='%d/%m/%Y'

conv_dates_series(test_df, "file2", old_date_format, new_date_format)


# ## First 5 rows of file 2 with the Date Created column changed to show DD/MM/YYYY

# In[17]:


file2.head(5)


# # Combine columns 'Street Number' and 'Street' into one 'Address' column

# ## Read csv from excel document file 3 and show first 5 columns

# In[18]:


file3 = pd.read_csv('../Keap/File 3.csv')
file3.head(5)


# ## Combine columns and change the datatype temporary to string

# In[19]:


file3['Address'] = file3['Street Number'].apply(str) + ' ' + file3['Street'].apply(str)


# ## Show first 5 columns to demonstrate verification of combined columns

# In[20]:


file3.head(5)


# # Add 'Date Created' from file *2* and 'Address' column from file 3 into file 1 and make sure they are assigned to the correct record based on the Id

# ## Merged File 2 into File 1 joining by Id (Date Created was the only other column)

# In[ ]:


merged_data = file1.merge(file2,on=["Id"],) 
merged_data.head()


# ## Used previous merge file to merge file 3 joining by Id and adding Address Column

# In[ ]:


file1 = pd.merge(merged_data,file3[['Id','Address']],on='Id', how='left')
file1.head()

