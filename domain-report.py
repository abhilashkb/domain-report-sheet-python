import openpyxl
from openpyxl import cell
import whois
import dns

domainfile=openpyxl.load_workbook("domain-report.xlsx")

domains=domainfile["Sheet1"]

for i in range(2, domains.max_row + 1):
    domain=domains.cell(i,1).value
    print(domain)