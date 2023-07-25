import hashlib
import PyPDF2
import re
import pandas as pd
from ics import Calendar, Event, alarm, DisplayAlarm
import pytz
from datetime import datetime, timedelta

def create_exam_calendar(exam_data):
    for semester in range(1, 9):
        semester_events = [exam for exam in exam_data if exam['Semester'] == str(semester)]
        # print(semester_events)
        if not semester_events:
            continue
        
        #create calendar
        c = Calendar()
        
        # Add the X-WR-TIMEZONE property using the property method
        # c.property('X-WR-TIMEZONE', 'Europe/Athens')
        
        # Create a time zone object using pytz
        timezone = pytz.timezone('Europe/Athens')


        
        file_name = f"semester_{semester}_exams.ics"
        for event_data in semester_events:
            event = Event()
            # print("event_data['Lesson Name']: ",event_data['Lesson Name'])
            event.name = event_data['Lesson Name']
            # print("event_data['Exam Date']: ",event_data['Exam Date'])
            # exam_date = datetime.strptime(event_data['Exam Date'], '%A , %d/%m/%Y').date()
            exam_date = datetime.strptime(event_data['Exam Date'].split(',')[1].strip(), '%d/%m/%Y').date()

            time_range = event_data['Exam Time'].split(' - ')
            event.begin = timezone.localize(datetime.combine(exam_date, datetime.strptime(time_range[0].strip(), '%H:%M').time()))
            event.end = timezone.localize(datetime.combine(exam_date, datetime.strptime(time_range[1].strip(), '%H:%M').time()))
            event.location = event_data['Classrooms']
            event.description = f"Εξάμηνο: {event_data['Semester']}\n" \
                                f"Τμήμα: {event_data['Lesson Class']}\n" \
                                f"Κωδικός μαθήματος: {event_data['Lesson Code']}"
                                
            # Add a notification 30 minutes before the event (Nevermind, notifications dont work for some reason in Google Calendar)
            # event.alarms = [DisplayAlarm(trigger=timedelta(days=0, hours=-1, minutes=0, seconds=0),display_text=event_data['Lesson Name'])]
            
            event_uid = hashlib.md5(f"{event_data['Lesson Name']}-{semester}".encode()).hexdigest()
            event.uid= f"{event_uid}@unipi.com"
            
            c.events.add(event)

        with open(file_name, 'w', encoding='utf-8') as f:
            f.writelines(c)
        print(f"Calendar file '{file_name}' created successfully!")

def delete_unnecessary_text(original_text, lesson_code):
    print("original original_text: "+original_text)
    clean_lesson_name=original_text.replace("Eπιλογή", "")
    clean_lesson_name=clean_lesson_name.replace(lesson_code+"-", "")
    clean_lesson_name = re.sub(r'\s+', ' ', clean_lesson_name).strip() #to remove multiple whitespaces from lesson name
    
    # original_text=original_text.strip()
    # if original_text[:4]=='ΨΣ -':
        # return original_text[8:].strip()
    # elif " ΠΔΙ -" in original_text:
        # return original_text[22:].strip()
    # elif "-1-" in original_text:
        # return original_text[18:].strip()
    # elif original_text[:7]=='Eπιλογή':
        # return original_text[16:].strip()
    return clean_lesson_name

def extract_exam_data(pdf_file_path):
    exam_data = []
    ignore_text = [
        'ΕΛΛΗΝΙΚΗ  ΔΗΜΟΚΡΑΤΙΑ',
        'ΠΑΝΕΠΙΣΤΗΜΙΟ  ΠΕΙΡΑΙΩΣ',
        '2022-2023',
        'ΕΞΕΤΑΣΕΙΣ  ',
        'ΣΧΟΛΗ  ΤΕΧΝΟΛΟΓΙΩΝ  ΠΛΗΡΟΦΟΡΙΚΗΣ  ΚΑΙ  ΕΠΙΚΟΙΝΩΝΙΩΝ',
        'ΤΜΗΜΑ  ΨΗΦΙΑΚΩΝ  ΣΥΣΤΗΜΑΤΩΝ',
        # 'Διεύθυνση  Σπουδών     - 1 -ΩΡΕΣ ΑΙΘΟΥΣΕΣ ΕΞ .- ΤΥΠΟΣ ΜΑΘΗΜΑ ΤΜΗΜΑ ΚΩΔΙΚΟΣ',
        'Διεύθυνση  Σπουδών     - ',
        'ΤΜΗΜΑ  ΠΛΗΡΟΦΟΡΙΚΗΣ',
        # 'Πειραιάς :18/5/2023 ',
        'Πειραιάς :',
        'Παρατηρήσεις :',
        '1. Όπου  ΓΛ  αίθουσα  στο  κτίριο  Γρ . Λαμπράκη  21 & Διστόμου  και  όπου  Νκ  αίθουσα  στο  κτίριο  Δεληγιώργη  και  Τσαμαδού .',
        '2. Για  να  γίνουν  δεκτοί  οι φοιτητές',
        '2. Για  να  γίνουν  δεκτοί  οι  φοιτητές  στις  εξετάσεις  θα  πρέπει  :',
        'α) Να  φέρουν  μαζί τους  την  ακαδημαϊκή  τους  ταυτότητα .',
        'α) Να  φέρουν  μαζί  τους  την  ακαδημαϊκή  τους  ταυτότητα .',
        'α) Να φέρουν  μαζί τους την ακαδημαϊκή  τους ταυτότητα .',
        'β) Να  προσέλθουν  στη  σειρά  τους  σύμφωνα  με  το  πρόγραμμα  των  εξετάσεων .',
        'Ο ΑΝΤΙΠΡΥΤΑΝΗΣ Ο ΠΡΟΕΔΡΟΣ',
        'ΑΝ . ΚΑΘΗΓΗΤΗΣ  ΣΠΥΡΙΔΩΝ  ΡΟΥΚΑΝΑΣ ΚΑΘΗΓΗΤΗΣ  ΓΕΩΡΓΙΟΣ  ΕΥΘΥΜΟΓΛΟΥ',
        'ΑΝ . ΚΑΘΗΓΗΤΗΣ  ΣΠΥΡΙΔΩΝ  ΡΟΥΚΑΝΑΣ ΚΑΘΗΓΗΤΡΙΑ  ΜΑΡΙΑ  ΒΙΡΒΟΥ'
    ]

    with open(pdf_file_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        # page = reader.pages[0]
        # all_pages = []
        
        # for page_num in range(len(reader.pages)):
        #     # Get the page object
        #     page = reader.pages[page_num]
        #     # Extract the text content of the page
        #     # content = page.extract_text()
        #     content = page.extract_text().split('\n')
        #     # Append the content to the list
        #     all_pages.append(content)
        content=''
        for page_num in range(len(reader.pages)):
            # Get the page object
            page = reader.pages[page_num]
            # Extract the text content of the page
            # content = page.extract_text()
            if(content==''):
                content = page.extract_text()
            else:    
                content = content+'\n'+page.extract_text()
        
        
        # content = all_pages.extract_text()
        # lines=all_pages
        lines = content.split('\n')
        
        # lines = lines.split('\n')
        # print(lines)

        lessonFound=False
        dayFound=False
        newlineFound=False
        extendedDescriptionFound=False
        classrooms=''
        lesson_name=''
        rest_of_line=''
        name_day=''
        exam_time=''
        semester=''
        lesson_code=''
        for line in lines:
            print ("line: "+line)
            if any(ignore in line for ignore in ignore_text):
                print("line ignored: "+line)
                continue
            
            
            if 'Α  ' in line:
                line = line.replace('Α  ', 'Α ')

            
            match = re.search(r'(\w+\s+,\s+\d{2}/\d{2}/\d{4})', line)
            if match:
                name_day = match.group(1)
                print("nameday: "+name_day)
                dayFound=True
                continue
            # else:
                # dayFound=False
            
            if not extendedDescriptionFound:
                if dayFound:
                    
                    lessonFound=True
                    
                    if lessonFound and not newlineFound:
                        exam_time=line[:14]
                        line =line[14:].strip()
                        print("exam time: "+exam_time)
                        # print("rest of line: "+rest_of_line)
                        # print(rest_of_line[-1:])
                        # if(rest_of_line[-1:]==',' or rest_of_line[-1:]=='.'):
                        line=line.strip()
                        # print("rest_of_line[-2:]: "+rest_of_line[-2:]=='Π')
                        # if(line[-1:]==',' or line[-2]=='Π'):
                        matchManyClassrooms=r"\d,$|Π.$|φ\s,$" #matches "1,EOL" or "Π,EOL" or φ,EOL"
                        if(re.search(matchManyClassrooms, line[-3:])):
                            classrooms=line
                            newlineFound=True
                            continue
                        print("newlineFound: "+str(newlineFound))
                        # parts = re.split(r'\s(?=\d)', line, maxsplit=1)
                        
                        #case where classrooms are only in one line
                        # parts = re.split(r'(\s\d\s)', line)
                        # print("parts:")
                        # print(parts)
                        # classrooms=parts[0]
                        # semester=parts[1].strip()
                        # line=parts[2].strip()
                        # lessonFound=True
                        
                    # parts=re.split(r"\d\sΨΣ\s-[0-9]+-(?=\s)",rest_of_line, maxsplit=1)
                    # parts = re.split(r'\b\w+\s-\d+\b', rest_of_line, maxsplit=1)


                    if newlineFound:
                        # parts = re.split(r'\s(?=\d)', rest_of_line, maxsplit=1)
                        line=line.strip()
                        # print("line[-1]:",line[-1]+"|")
                        # print("line[-2]:",line[-2]+"|")
                        # print("line[-3]:",line[-3]+"|")
                        if(line[-1:]==','):
                            classrooms=classrooms+' '+ line + ' '
                            continue
                        # parts = re.split(r'\s+(\d+\s+ΨΣ\s+-\d+-\s+.+)$', line)
                        # parts = re.split(r'\s+(?=\d)', line, maxsplit=1)
                        # parts = re.split(r'(\s\d\s)', line)


                        # print("parts 1:")
                        # print(parts)
                        
                        # classrooms.join(parts[0])
                        
                        # classrooms=classrooms+parts[0]
                        # semester=parts[1]
                        # line=parts[2]
                        
                        # newlineFound=False
                    
                    parts = re.split(r'(\s\d\s)', line)
                    print("parts 2:")
                    print(parts)
                    # if not newlineFound:
                        # classrooms=parts[0]
                    # semester=parts[1].strip()
                    # line=parts[2].strip()
                    
                    classrooms=classrooms+parts[0]
                    semester=parts[1]
                    line=parts[2]
                    
                    lessonFound=True
                    
                    # print(parts)
                    # classrooms=parts
                    # rest_of_line="ΨΣ -"+parts[1]
                    # rest_of_line=parts[1].strip()
                    print("classrooms: " + classrooms)
                    # print("rest of line: "+rest_of_line)
                    
                    # semester=line[:1]
                    # line=line[1:].strip()
                    print("semester: "+semester)
                    # print("rest_of_line: "+ rest_of_line)
                
                # print("rest_of_line[:-7:]: "+rest_of_line[-7:])
                    
                #searches for ΨΣ -123 at the end of line
                # extendedDescriptionMatch=re.search(r'ΨΣ\s-\d\d\d',rest_of_line[-7:])
                # if (not extendedDescriptionMatch and not rest_of_line[-3:]=='ΠΔΙ' and not rest_of_line[-1]=='1'):
                #     lesson_name=rest_of_line
                #     extendedDescriptionFound=True
                #     continue
                
            
                    
                    
            line=line.strip()
            # if not extendedDescriptionFound:
            #     if rest_of_line[-3:]=='ΠΔΙ':
            #         lesson_name= lesson_name +' '+ rest_of_line[:-18]
            #         rest_of_line=rest_of_line[-18:].strip()
            #     elif rest_of_line[-2:]=='-1':
            #         lesson_name= lesson_name +' '+ rest_of_line[:-15]
            #         rest_of_line=rest_of_line[-15:].strip()
            #     else:
            #         lesson_name= lesson_name +' '+ rest_of_line[:-13]
            #         rest_of_line=rest_of_line[-13:].strip()
            # else:
                # if line[-3:]=='ΠΔΙ':
                    # lesson_name= lesson_name +' '+ line[:-18]
                    # rest_of_line=line[-18:].strip()
                # else:
                    # lesson_name= lesson_name + ' ' +line[:-13]
                    # rest_of_line=line[-13:].strip()
            
            exam_class_split = r"([Α-Ω]\s-\s[Α-Ω])"  # Matches "Α - Ω" pattern
                
            #searches for "Α - Ω" in rest of line
            if not re.search(exam_class_split, line): #"Α - Ω" not found so whole line is lesson_name
                lesson_name= lesson_name + ' ' + line #appends to lesson name in case of many iterations
                extendedDescriptionFound=True
                continue #read next line, loops until "Α - Ω" pattern is found
            
            parts = re.split(exam_class_split, line)
            lesson_name= (lesson_name + ' ' + parts[0])
            
            exam_class=parts[1].strip()
            lesson_code= parts[2] #removes all whitespace

            
            
            lesson_name=delete_unnecessary_text(lesson_name, lesson_code)
            print("lesson name: "+lesson_name)
            # print("rest_of_line: "+rest_of_line)
            
            # exam_class=rest_of_line[:6].strip()
            # lesson_code=rest_of_line[6:].strip()
            print("exam class: "+exam_class)
            
            lesson_code= re.sub(r'\s', '', parts[2]) #removes all whitespace from lesson_code
            print("lesson code: "+ lesson_code)
            
            # classrooms, semester, lesson_info = re.split(r'\s', line, maxsplit=3)
            # lesson_name, lesson_class, lesson_code = re.split(r'\s{3,}|\s-\s', lesson_info)

            exam_data.append({
                'Exam Date': name_day,
                'Exam Time': exam_time.strip(),
                'Classrooms': classrooms.strip(),
                'Semester': semester.strip(),
                'Lesson Name': lesson_name.strip(),
                'Lesson Class': exam_class.strip(),
                'Lesson Code': lesson_code.strip()
            })
            
            lessonFound=False
            newlineFound=False
            extendedDescriptionFound=False
            classrooms=''
            lesson_name=''
            rest_of_line=''
            # name_day=''
            exam_time=''
            semester=''
            lesson_code=''
                    

    return exam_data

pdf_file_path = 'dates.pdf'
exam_data = extract_exam_data(pdf_file_path)
df = pd.DataFrame(exam_data)
df.to_excel('exam_data.xlsx', index=False)

create_exam_calendar(exam_data)