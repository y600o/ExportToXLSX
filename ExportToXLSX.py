# -*- coding: utf-8 -*-
import arcpy
import sys
import os

# Python 2/3 compatibility
try:
    unicode
except NameError:
    unicode = str

# Get paths
script_dir = os.path.dirname(os.path.abspath(__file__))
lib_dir = os.path.join(script_dir, 'lib')

# Add library paths
lib_paths = [
    os.path.join(lib_dir, 'openpyxl'),
    os.path.join(lib_dir, 'et_xmlfile'), 
    os.path.join(lib_dir, 'jdcal'),
    lib_dir
]

for path in lib_paths:
    if os.path.exists(path) and path not in sys.path:
        sys.path.insert(0, path)

# Import openpyxl
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font


def safe_unicode_str(obj):
    """Safely convert object to string, handling Unicode properly in Python 2.7"""
    if obj is None:
        return ""
    
    # Handle geoprocessing value objects specifically
    try:
        if hasattr(obj, 'value'):
            # This might be a geoprocessing value object
            value = obj.value
            # Now safely convert the value
            if sys.version_info[0] == 2:
                if isinstance(value, unicode):
                    return value.encode('utf-8')
                elif isinstance(value, str):
                    try:
                        # Try to decode as utf-8 first
                        return value.decode('utf-8').encode('utf-8')
                    except UnicodeDecodeError:
                        try:
                            # Try GBK for Chinese characters
                            return value.decode('gbk').encode('utf-8')
                        except UnicodeDecodeError:
                            return repr(value)
                else:
                    return str(value)
            else:
                return str(value)
    except Exception:
        pass
    
    # Handle regular objects
    try:
        if sys.version_info[0] == 2:
            # Python 2 - handle Unicode carefully
            if isinstance(obj, unicode):
                return obj.encode('utf-8')
            elif isinstance(obj, str):
                try:
                    # Try to decode then re-encode to ensure proper encoding
                    return obj.decode('utf-8').encode('utf-8')
                except UnicodeDecodeError:
                    try:
                        return obj.decode('gbk').encode('utf-8')
                    except UnicodeDecodeError:
                        return repr(obj)
            else:
                return str(obj)
        else:
            # Python 3 - strings are already Unicode
            return str(obj)
    except (UnicodeEncodeError, UnicodeDecodeError):
        try:
            return repr(obj)
        except:
            return "encoding_error"

def safe_field_name(field_obj):
    """Safely get field name, handling Unicode properly"""
    try:
        name = field_obj.name if hasattr(field_obj, 'name') else str(field_obj)
        return safe_unicode_str(name)
    except:
        return str(field_obj)

def safe_field_alias(field_obj):
    """Safely get field alias, handling Unicode properly"""
    try:
        alias = field_obj.alias if hasattr(field_obj, 'alias') else str(field_obj)
        return safe_unicode_str(alias)
    except:
        return str(field_obj)

from datetime import datetime


class clsField(object):
    """ Class to hold properties and behavior of the output fields
    """
    @property
    def alias(self):
        return self._field.aliasName

    @property
    def name(self):
        return self._field.name

    @property
    def domain(self):
        return self._field.domain

    @property
    def type(self):
        return self._field.type

    @property
    def length(self):
        return self._field.length

    def __init__(self, f, i, subtypes):
        """ Create the object from a describe field object
        """
        self.index = None
        self._field = f
        self.subtype_field = ''
        self.domain_desc = {}
        self.subtype_desc = {}
        self.index = i

        # Handle Python 2/3 compatibility for dictionary iteration
        if hasattr(subtypes, 'items'):
            subtypes_iter = subtypes.items()
        else:
            subtypes_iter = subtypes.iteritems()
            
        for st_key, st_val in subtypes_iter:
            if st_val['SubtypeField'] == f.name:
                self.subtype_desc[st_key] = st_val['Name']
                self.subtype_field = f.name
            
            # Handle Python 2/3 compatibility for nested dictionary iteration
            if hasattr(st_val['FieldValues'], 'items'):
                field_values_iter = st_val['FieldValues'].items()
            else:
                field_values_iter = st_val['FieldValues'].iteritems()
                
            for k, v in field_values_iter:
                if k == f.name:
                    if len(v) == 2:
                        if v[1]:
                            self.domain_desc[st_key]= v[1].codedValues
                            self.subtype_field = st_val['SubtypeField']

    def __repr__(self):
        """ Nice representation for debugging  """
        return '<clsfield object name={}, alias={}, domain_desc={}>'.format(self.name,
                                                                self.alias,
                                                                self.domain_desc)

    def updateValue(self, row, fields):
        """ Update value based on domain/subtypes """
        value = row[self.index]
        if self.subtype_field:
            subtype_val = row[fields.index(self.subtype_field)]
        else:
            subtype_val = 0

        if self.subtype_desc:
            value = self.subtype_desc[row[self.index]]

        if self.domain_desc:
            try:
                value = self.domain_desc[subtype_val][row[self.index]]
            except:
                pass # not all subtypes will have domain

        return value

def get_field_defs(in_table, use_domain_desc):
    desc = arcpy.Describe(in_table)

    subtypes ={}
    if use_domain_desc:
        subtypes = arcpy.da.ListSubtypes(in_table)

    fields = []
    for i, field in enumerate([f for f in desc.fields
                                if f.type in ["Date","Double","Guid",
                                              "Integer","OID","Single",
                                              "SmallInteger","String"]]):
        fields.append(clsField(field, i, subtypes))

    return fields



       
def table_to_excel(in_table, output, use_field_alias=False, use_domain_desc=False):
    fieldNames_forExcel = []
    wb = Workbook()
    ws = wb.active
    
    arcpy.AddMessage("Creating Excel file: " + safe_unicode_str(output))

    fields = get_field_defs(in_table, use_domain_desc)
    actual_field_names = [field.name for field in fields]
    
    checkedFields = arcpy.GetParameter(4)
    stringCheckedFields = []
    fieldNames_forExcel = []

    if len(checkedFields) > 0:
        checked_fields_safe = []
        
        if len(checkedFields) == 1:
            field_string = safe_unicode_str(checkedFields[0])
            if ";" in field_string:
                checked_fields_safe = [f.strip() for f in field_string.split(";") if f.strip()]
            else:
                checked_fields_safe = [field_string]
        else:
            for check in checkedFields:
                field_name = None
                try:
                    if hasattr(check, 'value'):
                        field_name = safe_unicode_str(check.value)
                    elif hasattr(check, 'valueAsText'):
                        field_name = safe_unicode_str(check.valueAsText)
                    else:
                        field_name = safe_unicode_str(check)
                except:
                    continue
                
                if field_name:
                    checked_fields_safe.append(field_name)
        
        # Match user-selected fields with actual field objects
        selected_fields = []
        for field in fields:
            safe_name = safe_field_name(field)
            field_alias = safe_field_alias(field)
            
            if safe_name in checked_fields_safe or field_alias in checked_fields_safe or field.name in checked_fields_safe:
                selected_fields.append(field)
                stringCheckedFields.append(field.name)
                
                if (use_field_alias == "true"):
                    fieldNames_forExcel.append(safe_field_alias(field))
                else:
                    fieldNames_forExcel.append(safe_field_name(field))
        
        fields = selected_fields
        for i, field in enumerate(fields):
            field.index = i
    else:
        stringCheckedFields = actual_field_names
        if use_field_alias == True:
            fieldNames_forExcel = [safe_field_alias(i) for i in fields]
        else:
            fieldNames_forExcel = [safe_field_name(i) for i in fields]

    inputDesc = arcpy.Describe(in_table)
    sheetName = arcpy.GetParameterAsText(5)
    
    # Handle worksheet title
    try:
        if sheetName == "":
            title_text = inputDesc.name
        else:
            title_text = sheetName
        
        if sys.version_info[0] == 2:
            if isinstance(title_text, str):
                try:
                    ws.title = title_text.decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        ws.title = title_text.decode('gbk')
                    except UnicodeDecodeError:
                        ws.title = u"Sheet1"
            elif isinstance(title_text, unicode):
                ws.title = title_text
            else:
                ws.title = unicode(str(title_text))
        else:
            ws.title = str(title_text)
    except:
        ws.title = u"Sheet1" if sys.version_info[0] == 2 else "Sheet1"
    
    ws.append(fieldNames_forExcel)
    
    # Format header row
    for excelField in list(ws)[0:1]:  # Only first row (header)
        for field in excelField:
            field.font = Font(bold=True)
            field.alignment = Alignment(horizontal='center')
            try:
                ws.column_dimensions[field.column_letter].width = 25
            except:
                pass
    # Write data rows
    with arcpy.da.SearchCursor(in_table, stringCheckedFields) as cursor:
        for row in cursor:
            dataRowList = []
            for col_index, value in enumerate(row):
                if (fields[col_index].domain_desc or fields[col_index].subtype_desc):
                    value = fields[col_index].updateValue(row, stringCheckedFields)
                
                if value is not None:
                    dataRowList.append(safe_unicode_str(value))
                else:
                    dataRowList.append("")
            ws.append(dataRowList)
    
    wb.save(output)

if __name__ == "__main__":
    # Parameter order: Input_Layer_or_Table, Output_XLSX_File, Use_field_alias_as_column_header, Use_domain_and_subtype_description, {Fields}, {Sheet_Name}
    input_table = arcpy.GetParameter(0)          # Input_Layer_or_Table
    output_file = arcpy.GetParameterAsText(1)    # Output_XLSX_File
    use_field_alias = arcpy.GetParameterAsText(2)  # Use_field_alias_as_column_header
    use_domain_desc = arcpy.GetParameterAsText(3)  # Use_domain_and_subtype_description
    # Note: Fields (parameter 4) and Sheet_Name (parameter 5) are handled inside table_to_excel function

    table_to_excel(input_table, output_file, use_field_alias, use_domain_desc)
