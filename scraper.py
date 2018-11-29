import urllib.request
import xlsxwriter

from bs4 import BeautifulSoup

BASE_URL = "https://bcorporation.com.au"

class Company:
	def __init__(self):
		self.name = None
		self.intro = None
		self.certified = None
		self.location = None
		self.sector = None
		self.site = None
		self.detail = None
		
	def __repr__(self):
		string = ""
		string += self.name + '\n'
		string += self.intro  + '\n'
		string += self.certified  + '\n'
		string += self.location  + '\n'
		string += self.sector  + '\n'
		string += self.site  + '\n'
		string += self.detail
		return string

def get_from_soup(soup, tag, c):
	a = soup.find(tag, c)
	if a:
		return a.get_text()
	else:
		return None
	
def get_company_from_url(sub_url):
	sub_page = urllib.request.urlopen(sub_url).read()
	sub_soup = BeautifulSoup(sub_page, 'lxml')
	company = Company()
	company.name = sub_soup.find('h1','serif heading3 mt-0').get_text()
	company.intro = get_from_soup(sub_soup, 'div', "field field-name-field-products-and-services field-type-text field-label-hidden")
	company.certified = get_from_soup(sub_soup, 'div', "field field-name-field-date-certified field-type-datestamp field-label-hidden sans-serif")
	company.location = get_from_soup(sub_soup, 'div', "field field-name-field-country field-type-text field-label-hidden sans-serif")
	company.sector = get_from_soup(sub_soup, 'div', "field field-name-field-sector field-type-text field-label-hidden sans-serif")
	company.site = get_from_soup(sub_soup, 'div', "field field-name-field-products-and-services field-type-text field-label-hidden")
	company.detail = get_from_soup(sub_soup, 'div', "field field-name-body field-type-text-with-summary field-label-hidden")
	return company

def write_to_xls(company_list):
	workbook = xlsxwriter.Workbook('Company_data.xlsx')
	worksheet = workbook.add_worksheet()

	# Start from the first cell. Rows and columns are zero indexed.
	row = 1

	# Iterate over the data and write it out row by row.
	for company in company_list:
		col = 0
		worksheet.write(row, col,     company.name)
		worksheet.write(row, col + 1, company.intro)
		worksheet.write(row, col + 2, company.certified)
		worksheet.write(row, col + 3, company.location)
		worksheet.write(row, col + 4, company.sector)
		worksheet.write(row, col + 5, company.site)
		worksheet.write(row, col + 6, company.detail)
		row += 1

	workbook.close()
	return


def main():
	verbose = True

	page = urllib.request.urlopen(BASE_URL + '/directory').read()
	soup = BeautifulSoup(page, 'lxml')
	company_list = list()

	for item in soup.find_all('article'):
		sub_url = BASE_URL + item.get('about')
		if verbose:
			print('Visiting ', sub_url)
		company_list.append(get_company_from_url(sub_url))

	for pagenum in range(1, 169):
		DIRURL = BASE_URL + '/directory?page=' + str(pagenum)
		page = urllib.request.urlopen(DIRURL).read()
		soup = BeautifulSoup(page, 'lxml')
		for item in soup.find_all('article'):
			sub_url = BASE_URL + item.get('about')
			if verbose:
				print('Visiting ', sub_url)
			company_list.append(get_company_from_url(sub_url))

	print(len(company_list), ' companies visited!')

	# Create a workbook and add a worksheet.
	write_to_xls(company_list)
	return 0


if __name__ == "__main__":
	main()

