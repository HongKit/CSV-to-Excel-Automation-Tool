################ Python automated test report generator ################

Version: 0.2
By: Hong-Kit Wong

Usage:
	
	Generate xlsx report using raw log files from the server

	Output file: Report Output/output_report

Please put the following documents in the designated folder:

	- Metadata/VMS_slots.csv

		- A csv file which the first ROW contains the slot numbers for VMS slots

		- Default values: [1,3,6,9,11,13,15]
	
	- Metadata/IPC_slots.csv

		- A csv file which the first ROW contains the slot numbers for IPC slots

		- Default values: [2,4,10,12,14,16]


	- Metadata/MKB_slots.csv

		- A csv file which the first ROW contains the slot numbers for MOCKINGBIRD slots

		- Default values: [6,7,8]

	- Metadata/Version_names.csv

		- A csv file which contains the APR version names and their corresponding slot numbers

		- Each line is one APR version

		- Line Format: VMS 4.0-8.05.05, [slot_numbers...]

		- Line Example: VMS 4.0-8.05.05,4,5,6,7,8,9,10,11,12

		- Default values: 
			apr_versions = ["VMS 4.0-8.05.05", "VMS-APR3.02-6.33.5", "VMS-3.1 - 6.55"]
			slots_in_apr = {
				"VMS 4.0-8.05.05": [4,5,6,7,8,9,10,11,12], 
				"VMS-APR3.02-6.33.5": [13,14,15,16], 
				"VMS-3.1 - 6.55": [1,2,3]
			}

	- Test Descriptions/{}

		- The Test Description folder contains descriptions to all Test Cases

		- Use Test Case ID as file name. E.G.: 1b.txt, 2.1a.txt

		- If file is not found for a test case, "No description specified" will be displayed on the spreadsheet

	- logs/logs{0}.csv

		- The raw log files from the server

		- The range of {0} MUST match the slot numbers defined above (in the metadata section)

		- Might use the tool Copy logs from server to automatically receive log files from the server
			- Must run in command line, type ./Copy_logs_from_server.exe for usage guide