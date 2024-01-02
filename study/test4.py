# https://plotly.com/blog/automate-excel-reports-with-python/

# 피벗 테이블(Pivot Tables)
           
# 또한 이 데이터를 그룹화, 필터링 또는 원하는 특정 기준에 따라 정렬함으로써 
# 대량의 데이터를 요약하고 분석할 수 있는 피벗 테이블을 생성할 수 있습니다. 

#Dataset used in this example:

#import the required packages
import pandas as pd
import plotly.express as px

# Load the Excel file into a pandas DataFrame
df = pd.read_excel("supermarket_sales4.xlsx")

# Create a pivot table
pivot_df = pd.pivot_table(df, values='Total', index='Gender', columns='Payment', aggfunc='sum')

# Export the pivot table to an Excel file
pivot_df.to_excel('output_file4.xlsx', sheet_name='Sheet1', index=True)
print(f"Excel 파일이 생성 되었습니다")

# --------------------------------------------------------------------------

# Additionally, you can create a Plotly figure for the same to visualize things graphically:

# Create a Plotly figure
fig = px.imshow(pivot_df)

# Show the figure
fig.show()
print(f"그림을 보여줍니다")

