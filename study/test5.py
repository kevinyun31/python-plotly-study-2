# https://plotly.com/blog/automate-excel-reports-with-python/

# 피벗 테이블(Pivot Tables)
           
# 또한 이 데이터를 그룹화, 필터링 또는 원하는 특정 기준에 따라 정렬함으로써 
# 대량의 데이터를 요약하고 분석할 수 있는 피벗 테이블을 생성할 수 있습니다. 

#Dataset used in this example:

#필요한 패키지 가져오기
import pandas as pd
import plotly.express as px

# Excel 파일을 pandas DataFrame에 로드합니다.
df = pd.read_excel("supermarket_sales5.xlsx")

# 피벗 테이블을 생성합니다
pivot_df = pd.pivot_table(df, values='Total', index='Gender', columns='Payment', aggfunc='sum')

# 피벗 테이블을 Excel 파일로 내보내기
pivot_df.to_excel('output_file5.xlsx', sheet_name='Sheet31', index=True)
print(f"Excel 파일이 생성 되었습니다")

# --------------------------------------------------------------------------

# 또한 동일한 내용을 그래픽으로 시각화하기 위해 Plotly 그림을 만들 수 있습니다.

# Plotly 그림을 만듭니다.
fig_bar = px.bar(pivot_df, x=pivot_df.index, y=pivot_df.columns, title='Payment Distribution by Gender')

# 그림을 보여주세요
fig_bar.show()
print(f"그래프를 보여줍니다")

