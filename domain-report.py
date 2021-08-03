import openpyxl
from openpyxl import cell
import whois
from dns import resolver

domainfile=openpyxl.load_workbook("domain-report.xlsx")

domains=domainfile["Sheet1"]

for i in range(2, domains.max_row + 1):
    domain=domains.cell(i,1).value
    print(domain)
    
    Arecord = resolver.resolve(domain)
    record=""
    for ipval in Arecord:
        print(ipval.to_text())
        record=record+str(ipval)+"\n"
        
    domains.cell(i,2).value=record
    MXrecord = resolver.resolve(domain,'mx')
    record=""
    for ipval in MXrecord:
        print(ipval.to_text())
        record=record+str(ipval)+"\n"
    domains.cell(i,3).value=record
    
    TXTrecord = resolver.resolve(domain,'txt')
    record=""
    for ipval in TXTrecord:
        print(ipval.to_text())
        record=record+str(ipval)+"\n"
    domains.cell(i,6).value=record
    
    dwhois=whois.whois(domain)
    #print(dwhois)
    #domains.cell(i,4).value=
    print(dwhois['expiration_date'][0].strftime('%m/%d/%Y'))
    domains.cell(i,5).value=str(dwhois['registrar'])
    



domainfile.save("solution.xlsx")    