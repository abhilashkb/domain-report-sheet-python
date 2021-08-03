import whois
import datetime
domain="google.com"
dwhois=whois.whois(domain)
print((dwhois['expiration_date'][1].strftime('%m/%d/%Y')))