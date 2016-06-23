################ Python automated test report generator ################

Version: 0.1
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

		- A csv file which the first ROW contains the APR Version names needed to generate the report

		- Default values: ["VMS 4.0-8.05.05", "VMS-APR3.02-6.33.5", "VMS-3.1 - 6.55"]

	- Test Descriptions/{}

		- The Test Description folder contains descriptions to all Test Cases

		- Use Test Case ID as file name. E.G.: 1b.txt, 2.1a.txt

		- If file is not found for a test case, "No description specified" will be displayed on the spreadsheet

	- logs/{0}/logs{1}.csv

		- The raw log files from the server

		- {0} MUST match the APR Version name defined above (in the metadata section)

		- The range of {1} MUST match the slot numbers defined above (in the metadata section)

		- Might use the tool Copy logs from server to automatically receive log files from the server