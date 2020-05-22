#!/home/joaofortunato/.virtualenvs/config_man_env/bin/python3.6

import os
import regex as re
import pandas as pd
import numpy as np


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
	drawing_parts_list = pd.read_excel(drawing_path,
	sheet_name = "DADOS PEÇAS", header = [20, 21])

	columns_index = drawing_parts_list.columns
	flat_columns_index = columns_index.to_flat_index()
	drawing_parts_list.columns = [e[1] if n < 6 else e[0] \
	for n, e in enumerate(flat_columns_index)]

	drawing_parts_list["NOTAS"] = \
	drawing_parts_list["NOTAS"].replace(to_replace = np.nan, value = "-")

	drawing_parts_list.dropna("columns", how = "all", inplace = True)

	# Add here: Function to deal with "-1 TO -29" type of qty columns.

	drawing_parts_list["DRAWING"] = drawing_path[-17:-5]

	drawing_parts_list = drawing_parts_list.melt(id_vars = ["REFERÊNCIA",
	"Nº DA PEÇA / \nESPECIF. DO MATERIAL", "DESIGNAÇÃO DA PEÇA OU MATERIAL",
	"EMPRESA/ORGANIZAÇÃO*", "NOTAS", "DRAWING"], var_name = "CONFIGURATION",
	value_name = "QUANTITY")
	drawing_parts_list["QUANTITY"] = \
	drawing_parts_list["QUANTITY"].replace(to_replace = " ", value = np.nan)

	drawing_parts_list.dropna("index", subset = ["QUANTITY"], inplace = True)
	drawing_parts_list["CONFIGURATION"] = drawing_parts_list["DRAWING"] +\
	drawing_parts_list["CONFIGURATION"].astype(str)

	return drawing_parts_list

def main():
	with os.scandir("Drawings/") as files:
		for file in files:
			if is_valid_drawing(file.name):
				# Scan file and pass xls to pandas dataframe
				path_to_file = os.curdir + "/Drawings/" + file.name
				drawing_dataframe = drawing_to_dataframe(path_to_file)

				db_drawings.append(drawing_dataframe, ignore_index = True)
if __name__ == "__main__":
	# path_to_file = os.curdir + "/Drawings/" + "15P18DR30000.xlsx"
	# pd.set_option("display.max_rows", None, "display.max_columns", None)
	# dataframe_test = drawing_to_dataframe(path_to_file)
	#
	# print(dataframe_test)
	# print(dataframe_test.columns)
	main()
