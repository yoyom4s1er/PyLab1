import pandas
import xlrd

excel_file_sheet_1 = pandas.read_excel("источник1.xls", "Ремонт")
excel_file_sheet_2 = pandas.read_excel("источник1.xls", "Mash")

excel_file_sheet_1['Цена ремонта'] = excel_file_sheet_1['Цена ремонта'].apply(lambda x: int(x) if isinstance(x,
int) else int(x.split()[0]))

excel_file_sheet_2.rename(columns={ excel_file_sheet_2.columns[3]: "Фирма-производитель", excel_file_sheet_2.columns[1]: "Паспорт"}, inplace=True)
excel_file_sheet_2['Паспорт'] = excel_file_sheet_2['Паспорт'].apply(lambda x: x.split('/')[0]+ "" + x.split('/')[1][:7])

pandas.set_option('display.max_columns', 500)

print(excel_file_sheet_1.to_string(), '\n')
print(excel_file_sheet_2.to_string(), '\n')

print('Таблица Мастер')
selected_columns_master = excel_file_sheet_1.iloc[:, [1]]
df_master = selected_columns_master.drop_duplicates()
df_master.insert(0, 'Id', range(1, len(df_master) + 1))
print(df_master.to_string(), '\n')

print('Таблица Вид операции')
selected_columns_operation = excel_file_sheet_1.iloc[:, [3]]
df_operation = selected_columns_operation.drop_duplicates()
df_operation.insert(0, 'Id', range(1, len(df_operation) + 1))
print(df_operation.to_string(), '\n')

print('Таблица Владелец')
df_owner = excel_file_sheet_2.iloc[:, [2, 1, 0]]
df_owner.insert(0, 'Id', range(1, len(df_owner) + 1))
print(df_owner.to_string(), '\n')

print('Таблица Фирма-производитель')
selected_columns_producer = excel_file_sheet_2.iloc[:, [3]]
df_producer = selected_columns_producer.drop_duplicates()
df_producer.insert(0, 'Id', range(1, len(df_producer) + 1))
print(df_producer.to_string(), '\n')

print('Таблица Марка')
selected_columns_mark = excel_file_sheet_2.iloc[:, [4,3]]
df_mark = selected_columns_mark.drop_duplicates()
df_merged = pandas.merge(df_mark, df_producer, on='Фирма-производитель', how='left')
df_merged.rename(columns={'Id': 'id_Фирма-производитель'}, inplace=True)
df_mark_final = df_merged[["Марка", "id_Фирма-производитель"]]
df_mark_final.insert(0, 'Id', range(1, len(df_mark_final) + 1))
print(df_mark_final.to_string(), '\n')


print('Таблица Машина')
selected_columns_car = excel_file_sheet_2.iloc[:, [6,5,4,1]]
df_car = selected_columns_car.drop_duplicates()
df_merged_car = pandas.merge(df_car, df_mark_final, on='Марка', how='left')
df_merged_car.rename(columns={'Id': 'id_Марка'}, inplace=True)
df_merged_car_2 = pandas.merge(df_merged_car, df_owner, on="Паспорт", how='left')
df_merged_car_2.rename(columns={'Id': 'id_Владелец'}, inplace=True)
df_car_final = df_merged_car_2[["ВИН", "Год", "id_Марка", "id_Владелец"]]
df_car_final.insert(0, 'Id', range(1, len(df_car_final) + 1))
print(df_car_final.to_string(), '\n')

print('Таблица Фактов')
df_merged_fact = pandas.merge(excel_file_sheet_1, df_master, on="Мастер", how="left")
df_merged_fact.rename(columns={'Id': 'id_мастер'}, inplace=True)
df_merged_fact = pandas.merge(df_merged_fact, df_operation, on="Операция", how="left")
df_merged_fact.rename(columns={'Id': 'id_операция'}, inplace=True)
df_merged_fact = pandas.merge(df_merged_fact, df_car_final, on="ВИН", how="left")
df_merged_fact.rename(columns={'Id': 'id_машина'}, inplace=True)
df_fact_final = df_merged_fact[["Дата", "id_мастер", "id_машина", "id_операция", "кол-часов", "Цена ремонта", "Коээффициент мастера"]]
print(df_fact_final.to_string(), '\n')