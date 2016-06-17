""" Assumptions """
""" No new/depricated tests between different versions """
""" Will be fixed in later versions"""


from openpyxl import load_workbook	# for writing the report
from openpyxl import Workbook
import csv 							# for reading from csv log files
import os, fnmatch					# for finding the filename of the Report file
from collections import Counter		# for counting number of times a test is run
from statistics import mean

##################### variables to tune based on future changes #####################
apr_versions = [3.0, 3.1, 4.0]
earliest_version = apr_versions[0]
latest_version = apr_versions[-1]
slot_numbers = [5,6]
IPC_slots = [5]
VMS_slots = [6]
##################### variables to tune based on future changes #####################

################################## util functions ###################################
def find_file_name_by_pattern(pattern, path):
    for root, dirs, files in os.walk(path):
        for name in files:
            if fnmatch.fnmatch(name, pattern):
                return os.path.join(root, name)

def read_csv(apr_version, slot_number, log_file_directories):
	try:
		log_file_contents = []
		log_file_name = '{}/APR{}/logs{}.csv'.format(log_file_directories, str(apr_version), str(slot_number))
		with open(log_file_name, newline='') as csvfile:
			spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
			for row in spamreader:
				if row[3] == "PASS":
					log_file_contents.append(row)
		return log_file_contents
	except OSError as err:
		print("\n\nThe file {} does not exists!!!\n\n".format(log_file_name))
		raise

def collect_from_csv(log_file_directories):
	# read and store date from log files
	IPC_records = {apr_version: [] for apr_version in apr_versions}
	VMS_records = {apr_version: [] for apr_version in apr_versions}
	for apr_version in apr_versions:
		for slot_number in slot_numbers:
			log_file_contents = read_csv(apr_version, slot_number, log_file_directories)
			if slot_number in IPC_slots:
				IPC_records[apr_version].extend(log_file_contents)
			else: # VMS
				VMS_records[apr_version].extend(log_file_contents)
	# # check if the length of each list in each dict is the same
	# for apr_version in apr_versions:
	# 	if len(IPC_records[apr_version]) != len(VMS_records[apr_version]):
	# 		raise Exception("Error! Check if the number of test runs match for different apr versions!")
	# 	elif len(IPC_records[apr_version]) == 0:
	# 		raise Exception("Error! Check if the log files are empty!")
	return IPC_records, VMS_records

def pad(records, tc, tc_name, num_to_pad):
	selected_records = [record for record in records if record[0] == tc]
	avg = mean([float(selected_record[2]) for selected_record in selected_records])
	pad_record = [tc, tc_name, str(avg), "PADDED_DATA", "PADDED_DATA"]
	return records + num_to_pad * [pad_record]


def check_and_pad(IPC_records, VMS_records):
	for apr_version in apr_versions:
		r1 = IPC_records[apr_version]
		r2 = VMS_records[apr_version]
		counter1 = Counter([(elem[0], elem[1]) for elem in r1])
		counter2 = Counter([(elem[0], elem[1]) for elem in r2])
		for tc, tc_name in counter1:
			if counter1[(tc, tc_name)] < counter2[(tc, tc_name)]:
				IPC_records[apr_version] = pad(IPC_records[apr_version], tc, tc_name, counter2[(tc, tc_name)] - counter1[(tc, tc_name)])
			elif counter1[(tc, tc_name)] > counter2[(tc, tc_name)]:
				VMS_records[apr_version] = pad(VMS_records[apr_version], tc, tc_name, counter1[(tc, tc_name)] - counter2[(tc, tc_name)])
	return IPC_records, VMS_records

################################## util functions ###################################

################################### main function ###################################
def write_report(report_name_pattern = "PerformanceData_APR*.xlsx", report_directory = 'Report', \
				 log_file_directories = "logs"):
	"""
	usage: write data from csv log files into the report xlsx file
	parameters: 
		report_name_pattern: the fnmatch pattern for finding the report file
			default = "PerformanceData_APR*.xlsx"
		report_directory: the directory that holds the report xlsx file
			default = 'Report'
		log_file_directories: the directory that holds the log csv file
			default = "logs"
	"""
	# extract records from log files
	IPC_records, VMS_records = collect_from_csv(log_file_directories)
	# sort each list in each dict by the first column
	for apr_version in apr_versions:
		IPC_records[apr_version] = sorted(IPC_records[apr_version], key=lambda elem: elem[0])
		VMS_records[apr_version] = sorted(VMS_records[apr_version], key=lambda elem: elem[0])
	# Check data and pad if necessary
	IPC_records, VMS_records = check_and_pad(IPC_records, VMS_records)
	

	for apr_version in apr_versions:
		print(len(IPC_records[apr_version]), len(VMS_records[apr_version]))


	# write to the report file
	# report_file_name = find_file_name_by_pattern(report_name_pattern, report_directory)
	# wb = load_workbook(report_file_name)
	# sheet_ranges = wb['PerformanceDATA']
	# wb.save('document.xlsx')


################################### main function ###################################
if __name__ == '__main__':

    write_report()

	# IPC_records, VMS_records = collect_from_csv("logs")
	# print(IPC_records[3.0][:3])
	# print()
	# print(pad(IPC_records[3.0][:3], "1b", 'Measure Ten Channel Change Time', 10))
