#! /usr/bin/env python2

import argparse
import xlrd
import subprocess

VALID_TITLES = ["Year Of Study","Student Number","First Name","Last Name"]

def main(workbook_name,ldif_name):
	class_list = xlrd.open_workbook(workbook_name)
	rows = extract(class_list)

	rows = strip_unused_rows(rows)
	rows = strip_unused_cols(rows)

	rows = add_usernames(rows)
	rows = add_passwords(rows)

	if ldif_name == None:
		ldif_name = workbook_name[:-3] + "ldif"

	make_ldif(rows,ldif_name)
	#print rows
	
def extract(xl_file):
	""" goes through the xls file and extracts the user data """
	sh = xl_file.sheet_by_index(0)
	rows = []
	for row in range(sh.nrows):
		cols = []
		for col in range(sh.ncols):
			cols.append(sh.cell(row,col).value)
		rows.append(cols)
	return rows

def find_headers_row(rows):
	header_row = 0
	for row in rows:
		for title in VALID_TITLES:
			if title in row:
				return header_row
		header_row += 1

def strip_unused_cols(rows):
	valid_col_numbers = find_valid_col_numbers(rows)
	new_rows = []
	for row in rows:
		new_row = []
		for col_num in valid_col_numbers:
			new_row.append(row[col_num])
		new_rows.append(new_row)
	return new_rows

def find_valid_col_numbers(rows):
	titles = rows[find_headers_row(rows)]	
	valid_col_numbers = []	
	col_number = 0
	for title in titles:
		if title in VALID_TITLES:
			valid_col_numbers.append(col_number)
		col_number += 1
	return valid_col_numbers

def strip_unused_rows(rows):
	return rows[find_headers_row(rows):]

def add_usernames(rows):
	headers = rows[find_headers_row(rows)]
	headers.append("Username")
	rows[find_headers_row(rows)] = headers
	new_rows = []
	new_rows.append(headers)
	for row in rows[find_headers_row(rows) + 1:]:
		first_name = row[headers.index("First Name")] 
		last_name = row[headers.index("Last Name")]
		username = last_name.lower().replace(" ","") + first_name[0].lower()
		row.append(username)
		new_rows.append(row)
	return new_rows

def add_passwords(rows):
	headers = rows[find_headers_row(rows)]
	headers.append("NTPassword")
	headers.append("LMPassword")
	headers.append("plainTextPassword")
	new_rows = []
	new_rows.append(headers)
	for row in rows[find_headers_row(rows) + 1:]:
		student_number = row[headers.index("Student Number")]
		lm_password,nt_password = smb_encrypt(student_number)
		row.append(nt_password)
		row.append(lm_password)
		row.append(student_number)
		new_rows.append(row)
	return new_rows

def smb_encrypt(password):
	""" Calls an smbencrypt which comes with freeradius-utils on Ubuntu 
		to encrypt the password given in smbencrypt form
	 """
	smbencrypt_output = subprocess.check_output(["smbencrypt",password])
	lm_password = smbencrypt_output[0:32].strip()	
	nt_password = smbencrypt_output[32:].strip()	
	return lm_password,nt_password	

def make_ldif(rows, ldif_filename):
	headers = rows[find_headers_row(rows)]
	uidNumbers = [1000,2000,3000,4000]
	for row in rows[find_headers_row(rows) + 1:]:
		first_name = row[headers.index("First Name")]
		last_name = row[headers.index("Last Name")]
		username = row[headers.index("Username")]
		yos = row[headers.index("Year Of Study")]
		nt_password = row[headers.index("NTPassword")]
		lm_password = row[headers.index("LMPassword")]
		plain_password = row[headers.index("plainTextPassword")]
		uidNumber = uidNumbers[int(yos) - 1]
		uidNumbers[int(yos) - 1] += 1
		smbRid = uidNumber*4
		entry = "" 
		entry += "dn: uid=" + username + ",ou=ug,dc=eie,dc=wits,dc=ac,dc=za \n"
		entry += "objectClass: account \n"
		entry += "objectClass: posixAccount \n"
		entry += "objectClass: sambaSamAccount \n"
		entry += "cn: " + first_name + " " + last_name + "\n"
		entry += "uid: " + username + "\n"
		entry += "displayName: " + first_name + " " + last_name + "\n"
		entry += "uidNumber: " + str(uidNumber) + "\n"
		entry += "gidNumber: "# + gidNumber + "\n"
		entry += "homeDirectory: /home/ug/" + username + "\n"
		entry += "loginShell: /bin/bash \n"
		entry += "# 4*uid to get RID (the last number)\n"
		entry += "sambaSID: S-1-5-21-3949128619-541665055-2325163404-" + str(smbRid) + "\n"
		entry += "sambaAcctFlags: [U         ] \n"
		entry += "sambaNTPassword: " + nt_password + "\n"
		entry += "sambaLMPassword: " + lm_password + "\n"
		write_to_ldif(entry)

def write_to_ldif(ldif_entry):
	#ldif_file = open(ldif_filename,'w')

if __name__ == "__main__":
	parser = argparse.ArgumentParser(
		description="""
			Convert an xls file in a specific format to an ldif file used.
		""")

	parser.add_argument('-i', '--input', dest='xls_url',
			   help='the path to the xls filename')
	parser.add_argument('-o', '--output', dest='ldif_url',
			   help='the path to the output ldif filename')

	args = parser.parse_args()
	main(args.xls_url,args.ldif_url)
