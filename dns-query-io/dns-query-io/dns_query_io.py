import dns.resolver
import pandas as pd
from fuzzywuzzy import fuzz
from fuzzywuzzy import process

inputdata = list()
outputdata = list()
input_filepath = "../dns_record_compiled.xlsm"
contentdb_filepath = "../content_db_smartcare.xlsx"
output_match_filepath = "../fuzzmatch_result.xlsx"
output_nslookup_filepath = "../nslookup_result.xlsx"

def resolver(url,i):
	result = list()
	try:
		answers=dns.resolver.resolve(url,'CNAME')
		for res in answers:
			result.append(res.target)
		print ('resolved record ',i)
	except dns.resolver.NoAnswer:
		result.append('No Answer')
		print ('resolved record ',i)
	except dns.resolver.NXDOMAIN:
		result.append('Record Does Not Exist')
		print ('resolved record ',i)
	except dns.resolver.NoNameservers:
		result.append('No Resolve')
		print ('resolved record ',i)
	except dns.resolver.Timeout:
		result.append('Timeout')
		print ('resolved record ',i)
	return result
	
# Function to read excel file source
def xls_read(fp,colname):
	df = pd.read_excel(fp,sheet_name=0)
	ql = df[colname].tolist()
	return ql

# Function to write excel file (list only)
def xls_write(rl,fp):
	df = pd.DataFrame(rl)
	df.to_excel(fp, index=False)

idx = 0
inputdata = xls_read(input_filepath,'Domain')
db_content_data = xls_read(contentdb_filepath,'APP_NAME')

for domain in inputdata:
	idx += 1
	outputdata.append(resolver(domain,idx))

xls_write(outputdata,output_nslookup_filepath)

idx = 0
fuzzresults_ratio = list()
fuzzresults_content = list()
for target_url in inputdata:
	matchrate=0
	content_owner=""
	for db_url in db_content_data:
		if matchrate < fuzz.token_set_ratio(target_url,db_url):
			matchrate = fuzz.token_set_ratio(target_url,db_url)
			content_owner = db_url
	print ('fuzzy checked record ',idx)
	idx += 1
	fuzzresults_ratio.append(matchrate)
	fuzzresults_content.append(content_owner)


matchdata = zip(fuzzresults_content,fuzzresults_ratio)
xls_write(matchdata,output_match_filepath)

