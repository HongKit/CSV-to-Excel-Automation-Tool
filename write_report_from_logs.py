""" Assumptions """
""" No new/depricated tests between different versions """
""" Will be fixed in later versions"""


from openpyxl import load_workbook	# for writing the report
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment
import csv 							# for reading from csv log files
import os, fnmatch					# for finding the filename of the Report file
from collections import Counter		# for counting number of times a test is run
from statistics import mean

##################### variables to tune based on future changes #####################
apr_versions = ["VMS-APR3.02-6.33.5", "VMS-3.1 - 6.55", "VMS 4.0-8.05.05"]
earliest_version = apr_versions[0]
latest_version = apr_versions[-1]
slot_numbers = [5,6]
IPC_slots = [5]
VMS_slots = [6]
##################### variables to tune based on future changes #####################

################################## util functions ###################################
def char_range(c1, length):
    """Generates the characters from `c1` to `c2`, inclusive."""
    for c in range(ord(c1), ord(c1) + length):
    # for c in range(ord(c1), ord(c1) + ord(length)):
        yield chr(c)

def find_file_name_by_pattern(pattern, path):
    for root, dirs, files in os.walk(path):
        for name in files:
            if fnmatch.fnmatch(name, pattern):
                return os.path.join(root, name)

def read_csv(apr_version, slot_number, log_file_directories):
	try:
		log_file_contents = []
		log_file_name = '{}/{}/logs{}.csv'.format(log_file_directories, str(apr_version), str(slot_number))
		with open(log_file_name, newline='') as csvfile:
			spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
			for row in spamreader:
				if row[3] == "PASS":
					row[2] = float(row[2])
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
def write_report(report_name_pattern = "PerformanceData_APR*.xlsx", report_directory = 'Report\Input', \
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

	# write to the report file
	report_file_name = find_file_name_by_pattern(report_name_pattern, report_directory)
	wb = load_workbook(report_file_name)
	ws = wb['PerformanceDATA']
	wb.remove_sheet(ws)
	ws = wb.create_sheet(title = 'PerformanceDATA')
	
	####################################### How to make it more general? #######################################
	ord_C = ord("C")
	ord_J = ord("C") + len(apr_versions) + 2 * (len(apr_versions)-1)
	
	ws['A1'], ws['B1'] = "TC", "TC Name"

	for ver_count, apr_version in enumerate(apr_versions):
		ws["{}1".format(chr(ord_C + ver_count))] = apr_versions[ver_count]
		ws["{}1".format(chr(ord_J + ver_count))] = apr_versions[ver_count]
		if ver_count != len(apr_versions) - 1:
			ws["{}1".format(chr(ord_C + len(apr_versions) + 2 * ver_count))] = \
				"Difference\n({}-{})".format(apr_versions[ver_count], apr_versions[-1])
			ws["{}1".format(chr(ord_C + len(apr_versions) + 2 * ver_count + 1))] = \
				"%Change"
			ws["{}1".format(chr(ord_J + len(apr_versions) + 2 * ver_count))] = \
				"Difference\n({}-{})".format(apr_versions[ver_count], apr_versions[-1])
			ws["{}1".format(chr(ord_J + len(apr_versions) + 2 * ver_count + 1))] = \
				"%Change"
	####################################### How to make it more general? #######################################


	first_key = list(VMS_records.keys())[0]
	for i in range(len(VMS_records[first_key])):
		ws["{}{}".format("A", i+2)] = VMS_records[first_key][i][0]
		ws["{}{}".format("B", i+2)] = VMS_records[first_key][i][1]

	for i in range(len(VMS_records[first_key])):
		for ver_count, apr_version in enumerate(apr_versions):
			ws["{}{}".format(chr(ord_C + ver_count), i+2)] = VMS_records[apr_version][i][2]
			ws["{}{}".format(chr(ord_J + ver_count), i+2)] = IPC_records[apr_version][i][2]
			if ver_count != len(apr_versions) - 1:
				ws["{}{}".format(chr(ord_C + len(apr_versions) + 2 * ver_count), i+2)] = \
					"={0}{2} - {1}{2}".format(chr(ord_C + ver_count), chr(ord_C + len(apr_versions) - 1), i + 2)
				ws["{}{}".format(chr(ord_C + len(apr_versions) + 2 * ver_count + 1), i+2)] = \
					"={0}{2}/{1}{2}*100".format(chr(ord_C + len(apr_versions) + 2 * ver_count), chr(ord_C + len(apr_versions) - 1), i + 2)
				ws["{}{}".format(chr(ord_J + len(apr_versions) + 2 * ver_count), i+2)] = \
					"={0}{2} - {1}{2}".format(chr(ord_J + ver_count), chr(ord_J + len(apr_versions) - 1), i + 2)
				ws["{}{}".format(chr(ord_J + len(apr_versions) + 2 * ver_count + 1), i+2)] = \
					"={0}{2}/{1}{2}*100".format(chr(ord_J + len(apr_versions) + 2 * ver_count), chr(ord_J + len(apr_versions) - 1), i + 2)

	light_green_fill  = PatternFill(fill_type='solid', start_color='FFDCEDC8', end_color='FFDCEDC8')
	light_blue_fill  = PatternFill(fill_type='solid', start_color='FF4FC3F7', end_color='FF4FC3F7')
	light_red_fill  = PatternFill(fill_type='solid', start_color='FFFFEBEE', end_color='FFFFEBEE')
	dark_green_fill  = PatternFill(fill_type='solid', start_color='FF689F38', end_color='FF689F38')	
	red_fill  = PatternFill(fill_type='solid', start_color='FFF44336', end_color='FFF44336')

	red_zone_width = len(apr_versions) + 2 * (len(apr_versions)-1)

	ord_A = ord("A")
	red_zone_width = len(apr_versions) + 2 * (len(apr_versions)-1)
	for letter in char_range("A", red_zone_width+2):
		ws['{}1'.format(letter)].fill = dark_green_fill
		ws.column_dimensions[letter].width = 18
	for letter in char_range(chr(ord_A + red_zone_width + 2), red_zone_width):
		ws['{}1'.format(letter)].fill = red_fill
		ws.column_dimensions[letter].width = 18
	for i in range(2, ws.max_row+1):
		for letter in char_range("C", red_zone_width):
			ws['{}{}'.format(letter, i)].fill = light_green_fill
			ws['{}{}'.format(letter, i)].alignment = Alignment(horizontal="center", vertical="center")
			ws['{}{}'.format(letter, i)].number_format = '0.00'
		for letter in char_range(chr(ord_A + red_zone_width + 2), red_zone_width):
			ws['{}{}'.format(letter, i)].fill = light_red_fill
			ws['{}{}'.format(letter, i)].alignment = Alignment(horizontal="center", vertical="center")
			ws['{}{}'.format(letter, i)].number_format = '0.00'
		if i == 2 or ws["A{}".format(i)].value != ws["A{}".format(i-1)].value:
			ws['A{}'.format(i)].fill = light_blue_fill
			ws['B{}'.format(i)].fill = light_blue_fill

	ws.row_dimensions[1].height = 60

	ws['{}{}'.format("C", 2)].number_format = '0.00'
	wb.save('Report/output/{}'.format(report_file_name[len(report_directory)+1:]))



################################### main function ###################################
if __name__ == '__main__':
    write_report()