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
	drawing_parts_list.columns = [columns_id[n] if n < 6 else str(e[0]).strip() \
	for n, e in enumerate(flat_columns_index)]

	drawing_parts_list["NOTES"] = \
	drawing_parts_list["NOTES"].replace(to_replace = np.nan, value = "-")

	# Clean and organize dataframe
	drawing_parts_list.dropna("columns", how = "all", inplace = True)

	drawing_parts_list = unstack_configurations(drawing_parts_list)

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

	ordered_columns = ["DRAWING", "CONFIGURATION", "REF.",\
	"PART NUMBER", "DESCRIPTION", "ORGANIZATION", "NOTES", "QUANTITY"]

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

def unstack_configurations(drawing_dataframe):
	"""
	In some cases, for ease of visualization, multiple configurations are
	stacked in the same column. For example:
	Column title: "-2 TO -20", means that this column is applicable for each
	configuration from -2 to -20.
	This function will transform this kind of columns in multiple columns,
	one for each configuration.
	"""
	configuration_pattern = r"-[0-9]+(TO|&)-[0-9]+"

	for config in drawing_dataframe.columns[5:-1]:
		if re.search(configuration_pattern, config.replace(" ", "")):
			config_limits = re.findall("(?<=-)[0-9]+", config.replace(" ", ""))
			for n in range(int(config_limits[0]), int(config_limits[1])+1):
				column_name = "-" + str(n)
				drawing_dataframe[column_name] = drawing_dataframe[config]
			drawing_dataframe = drawing_dataframe.drop(columns = config)
	return drawing_dataframe

def main():
	db_drawings = pd.DataFrame()

	with os.scandir("Drawings/") as files:
		for file in files:
			if is_valid_drawing(file.name):
				path_to_file = os.curdir + "/Drawings/" + file.name

				drawing_dataframe = drawing_to_dataframe(path_to_file)
				db_drawings = db_drawings.append(drawing_dataframe,
				ignore_index = True)

	db_drawings = db_drawings.sort_values(by=["DRAWING", "CONFIGURATION",
	 "REF."], ignore_index = True)
	db_drawings.to_excel("items_DB.xlsx", index=False, sheet_name="Items DB")

if __name__ == "__main__":
	main()
