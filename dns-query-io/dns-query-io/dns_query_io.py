import dns.resolver
import pandas as pd

inputdata = list()
outputdata = list()
filepath = 'C://Users/adity/source/repos/dns-query-io/dns_record_compiled.xlsm'

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
	

def xls_read(fp):
	df = pd.read_excel(fp,sheet_name=0)
	ql = df['Domain'].tolist()
	return ql

def xls_write(rl):
	df = pd.DataFrame(rl)
	df.to_excel('C://Users/adity/source/repos/dns-query-io/cname-result.xlsx', index=False)

idx = 0
inputdata = xls_read(filepath)
for domain in inputdata:
	idx += 1
	outputdata.append(resolver(domain,idx))

xls_write(outputdata)