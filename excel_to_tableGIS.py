# -*- coding: cp936 -*-
#
#将excel表与要素属性表关联，并实现对应字段属性的自动填充
#
import arcpy
import xlrd
import os
if __name__=="__main__":
    feature_table=arcpy.GetParameterAsText(0)

    excel_sheet=arcpy.GetParameterAsText(1)
    excel_table_path=os.path.dirname(excel_sheet)
    excel_sheet_name=os.path.basename(excel_sheet).split("$")[0]
    arcpy.AddMessage(excel_sheet_name)
    arcpy.AddMessage(excel_table_path)

    path=os.path.dirname(feature_table)
    arcpy.env.workspace=path

    workbook = xlrd.open_workbook(excel_table_path)
    for i in workbook.sheet_names():
        if excel_sheet_name==i:
            table = workbook.sheet_by_name(i)#获取sheet表
            excel_table_header = table.row_values(0) #获取表头
            col = table.col_values(0) #字段匹配列

    matching_field=arcpy.GetParameterAsText(2)
    theFields = arcpy.ListFields(feature_table)
    fields_header = [] #全部字段名
    for Field in theFields:
        fields_header.append(Field.name)

    # 获取EXCEL中的匹配字段的数据
    index = excel_table_header.index(matching_field)
    col_values = table.col_values(index)
    # 匹配出连接字段
    for i in fields_header:
        if matching_field.upper()==i.upper():
            # 匹配出需要关联的字段
            for ii in fields_header:
                for j in excel_table_header:
                    if ii.upper() == j.upper() and j.upper()!=matching_field.upper(): #匹配关联字段
                        field_name = []
                        field_name.append(matching_field)  # 字段表头的名称
                        field_name.append(ii)

                        input_index=excel_table_header.index(j)
                        input_col_values = table.col_values(input_index)# 获取excel表中的关联字段数据
                        with arcpy.da.UpdateCursor(feature_table, field_name) as cursor:
                            # 将excel中的数据输入到关联字段中
                            for row in cursor:
                                for jj in col_values:
                                    if row[0]==jj:
                                         row[1]= input_col_values[col_values.index(jj)]
                                         cursor.updateRow(row)
