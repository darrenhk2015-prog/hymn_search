"""Probe churchofgod.org.hk hymn URL patterns."""
import re
import urllib.request

url = 'https://churchofgod.org.hk/hymns/'
req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
html = urllib.request.urlopen(req, timeout=20).read().decode('utf-8', errors='replace')
links = sorted(set(re.findall(r'href="(https://churchofgod\.org\.hk/hymns/[^"]+)"', html)))
print('count', len(links))
for u in links[:8]:
    print(u)
# sample slug pattern
for u in links:
    m = re.search(r'/hymns/(h\d+-\d+)', u, re.I)
    if m:
        print('slug', m.group(1))
        break
