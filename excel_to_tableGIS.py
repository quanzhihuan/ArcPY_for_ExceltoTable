# -*- coding: cp936 -*-
#
#��excel����Ҫ�����Ա��������ʵ�ֶ�Ӧ�ֶ����Ե��Զ����
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
            table = workbook.sheet_by_name(i)#��ȡsheet��
            excel_table_header = table.row_values(0) #��ȡ��ͷ
            col = table.col_values(0) #�ֶ�ƥ����

    matching_field=arcpy.GetParameterAsText(2)
    theFields = arcpy.ListFields(feature_table)
    fields_header = [] #ȫ���ֶ���
    for Field in theFields:
        fields_header.append(Field.name)

    # ��ȡEXCEL�е�ƥ���ֶε�����
    index = excel_table_header.index(matching_field)
    col_values = table.col_values(index)
    # ƥ��������ֶ�
    for i in fields_header:
        if matching_field.upper()==i.upper():
            # ƥ�����Ҫ�������ֶ�
            for ii in fields_header:
                for j in excel_table_header:
                    if ii.upper() == j.upper() and j.upper()!=matching_field.upper(): #ƥ������ֶ�
                        field_name = []
                        field_name.append(matching_field)  # �ֶα�ͷ������
                        field_name.append(ii)

                        input_index=excel_table_header.index(j)
                        input_col_values = table.col_values(input_index)# ��ȡexcel���еĹ����ֶ�����
                        with arcpy.da.UpdateCursor(feature_table, field_name) as cursor:
                            # ��excel�е��������뵽�����ֶ���
                            for row in cursor:
                                for jj in col_values:
                                    if row[0]==jj:
                                         row[1]= input_col_values[col_values.index(jj)]
                                         cursor.updateRow(row)
