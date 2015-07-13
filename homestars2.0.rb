require 'rubygems'
require 'watir-webdriver'
require 'headless'
require 'write_xlsx'

headless = Headless.new
headless.start
browser = Watir::Browser.new
browser2 = Watir::Browser.new
workbook = WriteXLSX.new('homestars_data.xlsx')

sheets = [] # holds different sheets in files
pages = [ # holds list of categories used 
    'http://homestars.com/on/toronto/alarm-systems',
    'http://homestars.com/on/toronto/electricians',
    'http://homestars.com/on/toronto/general-contractors',
    'http://homestars.com/on/toronto/paint-wallpaper-contractors',
    'http://homestars.com/on/toronto/windows-doors',
    'http://homestars.com/on/toronto/fire-water-damage-restoration',
    'http://homestars.com/on/toronto/plumbing'
]

#specifying format
format = workbook.add_format(:align => 'left')
header_format = workbook.add_format(:align => 'left', :bold => 1)
title_format = workbook.add_format(:align => 'center', :bold => 1)

#cycles through every category
sheet_index = 0
pages.each do |service_url|

    #create new worksheet and add header
    sheets[sheet_index] = workbook.add_worksheet(service_url[32..-1])
    sheets[sheet_index].write(
        1, 0,
        [
            'Name',
            'Phone',
            'Location',
            'Website',
            'Score',
            'Products',
            'Services',
            'Brands',
            'Styles',
            'Specialties',
            'Year',
            '# of Employees',
            'Return Policy',
            'Payment Method',
            'Licenses',
            'Memberships',
            'Liability',
            'Worker Compensation',
            'Bonded',
            'Written Contract Provided',
            'Warranty Terms',
            'Diplomas',
            'Project Minimum',
            'Project Rate'
        ], header_format
    )

    #cycle through 10 pages of the category
    row_num = 2
    page = 1
    while page <= 10 do
        puts "Page #{page}: #{service_url[32..-1]}"
        browser.goto("#{service_url}?page=#{page}")
        
        sheets[sheet_index].merge_range( 0, 0, 0, 23, browser.title.strip[0..-33],title_format ) if page == 1

        # unhides all phone numbers in the page
        browser.spans(text: /Telephone/).each do |button|
            button.click
        end

        # cycles through all listings in the page
        browser.sections(:class => 'company').each do |company|
            if company.p(:class => "phone").exists?
                puts name = company.h1.text.tr(',','')
                number = company.p(:class => "phone").text

                if company.span(:text => /Location/).exists?
                    location = company.span(:text => /Location/).text[2..-10]
                else
                    location = 'N/A'
                end

                if company.a(:class => 'go-to-company-www').exists?
                    website = company.a(:class => 'go-to-company-www').href
                else
                    website = 'N/A'
                end

                rank = company.span(:class => 'reputation-rank').text

                # set N/A as default value of these categories
                products = services = brands = styles = specialties = 'N/A'
                year = num_employee = rtrn_policy = pymnt_mthd = licenses = member = liability = worker_comp = bonded = contract_prov = warranty = diplomas = proj_min = proj_rate = 'N/A'
                
                # companies with logos visable have additional info
                if company.div(:class => 'logo').exists?
                    browser2.goto company.h1.a.href

                    info_index = 0
                    browser2.dts.each do |info|
                        case info.text.strip
                        when 'Products'
                            products = browser2.dds[info_index].text
                        when 'Services'
                            services = browser2.dds[info_index].text
                        when 'Brands'
                            brands = browser2.dds[info_index].text
                        when 'Styles'
                            styles = browser2.dds[info_index].text
                        when 'Specialties'
                            specialties = browser2.dds[info_index].text
                        when 'YEAR ESPABLISHED'
                            year = browser2.dds[info_index].text
                        when 'NUMBER OF EMPLOYEES'
                            num_employee = browser2.dds[info_index].text
                        when 'PAYMENT METHOD'
                            pymnt_mthd = browser2.dds[info_index].text
                        when 'LICENSES'
                            licenses = browser2.dds[info_index].text
                        when 'LIABILITY INSURANCE'
                            liability = browser2.dds[info_index].text
                        when 'WORKERS COMPENSATION'
                            worker_comp = browser2.dds[info_index].text
                        when 'BONDED'
                            bonded = browser2.dds[info_index].text
                        when 'WRITTEN CONTRACT PROVIDED'
                            contract_prov = browser2.dds[info_index].text
                        when 'MEMBERSHIPS'
                            member = browser2.dds[info_index].text
                        when 'DIPLOMAS'
                            diplomas = browser2.dds[info_index].text
                        when 'PROJECT MINIMUM'
                            proj_min = browser2.dds[info_index].text
                        when 'PROJECT RATE'
                            proj_rate = browser2.dds[info_index].text
                        when 'WARRANTY TERMS'
                            warranty = browser2.dds[info_index].text
                        when 'RETURN POLICY'
                            rtrn_policy = browser2.dds[info_index].text
                        end

                        info_index += 1
                    end
                end

                # add row with all data for the current listing
                sheets[sheet_index].write(
                    row_num, 0,
                    [
                        name,
                        number,
                        location,
                        website,
                        rank,
                        products,
                        services,
                        brands,
                        styles,
                        specialties,
                        year,
                        num_employee,
                        rtrn_policy,
                        pymnt_mthd,
                        licenses,
                        member,
                        liability,
                        worker_comp,
                        bonded,
                        contract_prov,
                        warranty,
                        diplomas,
                        proj_min,
                        proj_rate
                    ],format
                )
            end
            row_num += 1
        end
        page += 1
    end
    sheet_index += 1
end

workbook.close
browser2.close
browser.close
headless.destroy

