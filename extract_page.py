import fitz
doc = fitz.open('copy.pdf') #copy.pdf is the textbook
page_no=32 #page to be read
page = doc.loadPage(page_no) 
text = page.getText('text')

search_string="Defining Software" #heading to be searched for
next_string="Software Application Domains" #the heading that'll come after search_string

if(text.find(search_string) == -1): #search_string not found in page_no
    print("couldn't find"+search_string+" !")
else:
    start=text.find(search_string) 
    length=len(search_string)
    next=start+length
    if(text.find(next_string) != -1): #case where search_string and next_string are in the same page
        start_next=text.find(next_string)
        text_input=text[next:start_next] #extracting paragraph between search_string and next_string
        outF = open('extract_page.txt',"w") #writing the paragraph to a file
        outF.write(search_string)
        outF.write("\n")
        outF.write(text_input)
    else: #if the two headings aren't in the same page
        text_input=text[next:] #extract all text coming after search_string
        outF = open('extract_page.txt',"w") #writing it to the output file
        outF.write(search_string)
        outF.write("\n")
        outF.write(text_input)
        flag=0 #flag=0 => next_string not found yet
        while(flag==0): #will keep on iterating until next_string is found in some page
            page_no=page_no+1
            page = doc.loadPage(page_no) #reading next page of page_no
            text = page.getText('text')
            temp=text.find(next_string)
            if(temp != -1):
                text_input=text[:temp]
                flag=1
            else:
                text_input=text
            outF.write(text_input)


    
       
