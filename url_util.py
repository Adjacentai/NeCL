import base64
from urllib.parse import urlparse
import re

telegram_pattern = re.compile(r"https://t\.me/[\w]+")  

gis_pattern = re.compile(r"https://link\.2gis\.ru/\S+")

def decode_2gis_link(link):
    encoded_part = link.split('/')[-1]
    
    try:
        decoded_bytes = base64.urlsafe_b64decode(encoded_part)
        decoded_text = decoded_bytes.decode('utf-8')
        
        matches = re.findall(r'https://t\.me/[^\s\n]+', decoded_text)
        if matches:
            return matches[0]
            
    except Exception as e:
        return None

    return None

def check_telegram_link(decoded_link):
    return 't.me' in decoded_link or 'telegram.me' in decoded_link

def find_telegram_links(html):
    direct_links = telegram_pattern.findall(html)
    
    if direct_links:
        return direct_links

    links = gis_pattern.findall(html)
    telegram_links = []

    for link in links:
        decoded_link = decode_2gis_link(link)
        if decoded_link:
            telegram_links.append(decoded_link)

    return telegram_links
