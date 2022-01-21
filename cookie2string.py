firefox_cookie = '''{
	"Куки запроса": {
		"_ga_9E363E7BES": "GS1.1.1642676178.3.1.1642679955.60",
		"abp": "0",
		"buyer_laas_location": "637640",
		"buyer_location_id": "637640",
		"dfp_group": "5",
		"lastViewingTime": "1642679960840",
		"luri": "moskva",
		"showedStoryIds": "87-86-85-84-83-82-79-78-77-76-75-74-71-69-68-61-50",
		"sx": "H4sIAAAAAAAC/wTAQQ6CMBAF0Lv8tYtqp/87vY20A8EdkZIo4e6+EyTZujg7vdDooSmyd5XUmrqjnjhQkWLO23vwFU+tm4a+x/7bJ1uW+Kwj4YZAvdMeInOx6/oHAAD//6y9vGNbAAAA",
		"u": "2t4du0g6.qdwuhb.a5e8mudzyg80",
		"v": "1642679950"
	}
}'''
print('\'' + firefox_cookie.replace('\t','')\
      .replace('\n','')\
      .replace('{"Куки запроса\": {','')\
      .replace('}}','')\
      .replace(':','=')\
      .replace('\"','') + '\'')
