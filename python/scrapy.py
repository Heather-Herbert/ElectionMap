from requests import get
from requests.exceptions import RequestException
from contextlib import closing
from bs4 import BeautifulSoup
import csv
from array import array

def simple_get(url):
    """
    Attempts to get the content at `url` by making an HTTP GET request.
    If the content-type of response is some kind of HTML/XML, return the
    text content, otherwise return None.
    """
    try:
        with closing(get(url, stream=True)) as resp:
            if is_good_response(resp):
                return resp.content
            else:
                return None

    except RequestException as e:
        log_error('Error during requests to {0} : {1}'.format(url, str(e)))
        return None


def is_good_response(resp):
    """
    Returns True if the response seems to be HTML, False otherwise.
    """
    content_type = resp.headers['Content-Type'].lower()
    return (resp.status_code == 200 
            and content_type is not None 
            and content_type.find('html') > -1)


def log_error(e):
    """
    It is always a good idea to log errors. 
    This function just prints them, but you can
    make it do anything.
    """
    print(e)

baseURL = 'https://www.streetcheck.co.uk/postcode/'
with open('C:\\Users\\heath\\python\\postcodes.csv', 'r') as csvFile:
    reader = csv.reader(csvFile)
    for row in reader:
        response = simple_get(baseURL + row[0])
        items = []
        if response is not None:
            html = BeautifulSoup(response, 'html.parser')
            names = set()
            for li in html.select('td'):
                #for name in li.text.split('\n'):
                items.append(li)
        with open('Demographic_' + row[0] + '.csv', 'w') as csvWritterFile:
            writer = csv.writer(csvWritterFile)
            writer.writerows(items)
