import re
import urllib.request

xml = urllib.request.urlopen(
    urllib.request.Request(
        'https://churchofgod.org.hk/wp-sitemap-posts-hymns-1.xml',
        headers={'User-Agent': 'Mozilla/5.0'},
    )
).read().decode('utf-8', 'replace')
urls = re.findall(r'<loc>(https://churchofgod\.org\.hk/hymns/[^<]+)</loc>', xml)
print('urls', len(urls))
for u in urls[:5]:
    print(u)
codes = []
for u in urls:
    m = re.search(r'/hymns/(h\d+-\d+)', u, re.I)
    if m:
        codes.append(m.group(1).lower())
print('with code', len(codes))
