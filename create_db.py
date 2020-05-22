#!/home/joaofortunato/.virtualenvs/config_man_env/bin/python3.6

import os
import regex as re
import pandas as pd
import numpy as np
from openpyxl import load_workbook


def is_valid_drawing(file_name):
	"""
	Checks if file name is in a valid form for a Drawing.
	The valid file pattern depends from organization to organization.
	"""
	valid_drawing_pattern = r"^[0-9]{2}P[0-9]{2}DR[0-9]{5}\.xls+"

	if re.search(valid_drawing_pattern, file_name):
		return True
	else:
		return False

def drawing_to_dataframe(drawing_path):
	"""
	Uses pandas to open the xlsx file of the drawing and convert it
	to a pandas dataframe.
	The file format is specific to the organization.
	"""
	header_line = which_version_of_template(drawing_path)
	drawing_parts_list = pd.read_excel(drawing_path,
	sheet_name = "DADOS PEÇAS", header = [header_line, header_line+1])

	# Standardize columns titles and flatten columns
	columns_index = drawing_parts_list.columns
	flat_columns_index = columns_index.to_flat_index()
	columns_id = ["AUX","REF.", "PART NUMBER", "DESCRIPTION", "ORGANIZATION",\
	 "NOTES"]
	drawing_parts_list.columns = [columns_id[n] if n < 6 else e[0] \
	for n, e in enumerate(flat_columns_index)]

	drawing_parts_list["NOTES"] = \
	drawing_parts_list["NOTES"].replace(to_replace = np.nan, value = "-")

	# Clean and organize dataframe
	drawing_parts_list.dropna("columns", how = "all", inplace = True)

	# Add here: Function to deal with "-1 TO -29" type of qty columns.

	drawing_parts_list["DRAWING"] = drawing_path[-17:-5]

	drawing_parts_list = drawing_parts_list.melt(id_vars = ["REF.",
	"PART NUMBER", "DESCRIPTION",
	"ORGANIZATION", "NOTES", "DRAWING"], var_name = "CONFIGURATION",
	value_name = "QUANTITY")

	drawing_parts_list["QUANTITY"] = \
	drawing_parts_list["QUANTITY"].replace(to_replace = " ", value = np.nan)
	drawing_parts_list.dropna("index", subset = ["QUANTITY"], inplace = True)

	drawing_parts_list["CONFIGURATION"] = drawing_parts_list["DRAWING"] +\
	drawing_parts_list["CONFIGURATION"].astype(str)

	ordered_columns = ["DRAWING", "CONFIGURATION", "REF.", "PART NUMBER",\
	"DESCRIPTION", "ORGANIZATION", "NOTES", "QUANTITY"]

	drawing_parts_list = drawing_parts_list[ordered_columns]

	return drawing_parts_list

def which_version_of_template(drawing_path):
	"""
	For this organization there 2 types of drawing templates being used.
	Each different version of the template has a different initial line where
	the parts list begins to be listed.
	This function access the .xlsx file and inspects a certain cell to get the
	true initial line.
	"""
	drawing_workbook = load_workbook(filename=drawing_path, read_only = True)
	parts_sheet = drawing_workbook["DADOS PEÇAS"]

	if parts_sheet["B13"].value == "REFERÊNCIA":
		initial_line = 11
	else:
		initial_line = 20

	return initial_line

def main():
	with os.scandir("Drawings/") as files:
		for file in files:
			if is_valid_drawing(file.name):
				# Scan file and pass xls to pandas dataframe
				path_to_file = os.curdir + "/Drawings/" + file.name
				drawing_dataframe = drawing_to_dataframe(path_to_file)

				db_drawings.append(drawing_dataframe, ignore_index = True)
if __name__ == "__main__":
	path_to_file = os.curdir + "/Drawings/" + "15P18DR30000.xlsx"
	pd.set_option("display.max_rows", None, "display.max_columns", None)
	dataframe_test = drawing_to_dataframe(path_to_file)

	print(dataframe_test)
	print(dataframe_test.columns)
	#main()
