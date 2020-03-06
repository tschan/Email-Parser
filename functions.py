def getTablesFromHTML(file, encoding):
    from bs4 import BeautifulSoup
    # open htm-File of email as string
    with open(file, "r", encoding=encoding) as f:
        xhtml = f.read()

    # parse string for tables
    soup = BeautifulSoup(xhtml, 'html.parser')
    # Sehr schöne List Comprehensions, sehr pythonic! :)
    tables = [
        [
            [td.get_text(strip=True) for td in tr.find_all('td')]
            for tr in table.find_all('tr')
        ]
        for table in soup.find_all('table')
    ]
    return tables
