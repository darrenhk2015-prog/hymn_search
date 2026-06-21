import re
import urllib.request

url = 'https://churchofgod.org.hk/hymns/page/2/'
h = urllib.request.urlopen(
    urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
).read().decode('utf-8', 'replace')
links = sorted(set(re.findall(r'https://churchofgod\.org\.hk/hymns/[^"\s<>]+', h)))
print('links', len(links))
for u in links[:10]:
    print(u)
