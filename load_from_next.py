#coding:utf-8
import lxml.html
import string
import codecs
from selenium import webdriver
f=codecs.open('test_rikunabinext_bussiness.txt','w','utf-8')

total_url='https://next.rikunabi.com/eigyo/lst_jb0100000000/'
pjs_path = 'C:/Users/tai.RD-WOODS/phantomjs-2.1.1-windows/bin/phantomjs'
dcap = {
    'marionette' : True
}
driver = webdriver.PhantomJS(executable_path=pjs_path, desired_capabilities=dcap)

driver.get(total_url)
total_root = lxml.html.fromstring(driver.page_source)
element=total_root.xpath('//p[@class="rnn-textM"]/span[1]')[0].text_content()
print(element)
for num in range(1,int(element),50):
	print(num)
	list_url=total_url+'crn'+str(num)+'.html'
	driver.get(list_url)

	l_root = lxml.html.fromstring(driver.page_source)
	for x in l_root.xpath('//p[@class="rnn-offerCatch__contents__title js-abScreen__main__text"]'):
		f.write(x.text_content().replace(' ','').replace('\n','').replace('■','\n').replace('★','\n').replace('□','').replace('◎','\n').replace('※','\n').replace('◆','\n').replace('●','\n').replace('【','').replace('】','。').replace('》','。').replace('《','').replace('☆','').replace('◇',''))
		f.write('\n')
f.close()