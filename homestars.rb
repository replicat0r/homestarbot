require 'rubygems'
require 'watir-webdriver'
require 'headless'

headless = Headless.new
headless.start
browser = Watir::Browser.new

browser.goto("http://homestars.com/on/toronto/categories")
pages = []
links = browser.div(:class => "categories_alphas").links.each{|link| pages << link.href}

pages.each do |service_url|
  page = 1
  while page <= 10 do
      puts "Page #{page}: #{service_url[32..-1]}"
      browser.goto("#{service_url}?page=#{page}")
      File.open("#{browser.url[32..-8]}.csv", 'w'){|f| f.puts 'Name,Phone,Location,Website,Score'} if page==1
      browser.spans(text: /Telephone/).each do |button|
        button.click
      end
      browser.sections(:class => 'company').each do |company|
        if company.p(:class => "phone").exists?
          name = company.h1.text.tr(',','')
          number = company.p(:class => "phone").text
          data = name+','+number
          if company.span(:text => /Location/).exists?
            data << ','+company.span(:text => /Location/).text[2..-10]
          else
            data << ',N/A'
          end

          if company.a(:class => 'go-to-company-www').exists?
            data << ','+company.a(:class => 'go-to-company-www').href
          else
            data << ',N/A'
          end
          data << ','+ company.span(:class => 'reputation-rank').text
          puts data
          open("#{service_url[32..-1]}.csv", 'a') { |f| f.puts data }
        end
      end
      page += 1
    end
  end



  browser.close
  headless.destroy
