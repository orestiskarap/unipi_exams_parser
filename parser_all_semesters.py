import hashlib
import PyPDF2
import re
import pandas as pd
from ics import Calendar, Event
import pytz
from datetime import datetime, timedelta

def create_exam_calendar(exam_data):
    for semester in range(1, 9):
        semester_events = [exam for exam in exam_data if exam['Semester'] == str(semester)]
        if not semester_events:
            continue
        
        #create calendar
        c = Calendar()
        
        # Create a time zone object using pytz
        timezone = pytz.timezone('Europe/Athens')


        
        file_name = f"semester_{semester}_exams.ics"
        for event_data in semester_events:
            event = Event()
            event.name = event_data['Lesson Name']
            exam_date = datetime.strptime(event_data['Exam Date'].split(',')[1].strip(), '%d/%m/%Y').date()

            time_range = event_data['Exam Time'].split(' - ')
            event.begin = timezone.localize(datetime.combine(exam_date, datetime.strptime(time_range[0].strip(), '%H:%M').time()))
            event.end = timezone.localize(datetime.combine(exam_date, datetime.strptime(time_range[1].strip(), '%H:%M').time()))
            event.location = event_data['Classrooms']
            event.description = f"Εξάμηνο: {event_data['Semester']}\n" \
                                f"Τμήμα: {event_data['Lesson Class']}\n" \
                                f"Κωδικός μαθήματος: {event_data['Lesson Code']}"
                                
            event_uid = hashlib.md5(f"{event_data['Lesson Name']}-{semester}".encode()).hexdigest()
            event.uid= f"{event_uid}@unipi.com"
            
            c.events.add(event)

        with open(file_name, 'w', encoding='utf-8') as f:
            f.writelines(c)
        print(f"Calendar file '{file_name}' created successfully!")

def delete_unnecessary_text(original_text, lesson_code):
    clean_lesson_name=original_text.replace("Eπιλογή", "") # to remove "Επιλογή"
    clean_lesson_name=clean_lesson_name.replace(lesson_code, "").strip() #to remove lesson code from name
    clean_lesson_name=clean_lesson_name[1:] #to remove dash seperating lesson code fron lesson name
    clean_lesson_name = re.sub(r'\s+', ' ', clean_lesson_name).strip() #to remove multiple whitespaces from lesson name
    
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
        content=''
        for page_num in range(len(reader.pages)):
            # Get the page object
            page = reader.pages[page_num]
            # Extract the text content of the page
            if(content==''):
                content = page.extract_text()
            else:    
                content = content+'\n'+page.extract_text()
        
        
        lines = content.split('\n')
        
        lessonFound=False
        dayFound=False
        newlineFound=False
        extendedDescriptionFound=False
        classrooms=''
        lesson_name=''
        name_day=''
        exam_time=''
        semester=''
        lesson_code=''
        for line in lines:
            # print ("line: "+line)
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
            
            if not extendedDescriptionFound:
                if dayFound:
                    
                    lessonFound=True
                    
                    if lessonFound and not newlineFound:
                        exam_time=line[:14]
                        line =line[14:].strip()
                        print("\nexam time: "+exam_time)
                        line=line.strip()
                        matchManyClassrooms=r"\d,$|Π.$|φ\s,$" #matches "1,EOL" or "Π,EOL" or φ,EOL"
                        if(re.search(matchManyClassrooms, line[-3:])): #check if classrooms are spanning multiple lines
                            classrooms=line
                            newlineFound=True
                            continue

                    if newlineFound: #in loop adding new line
                        line=line.strip()
                        # if(line[-1:]==','):
                        if(re.search(matchManyClassrooms, line[-3:])): #check if classrooms are spanning multiple lines
                            classrooms=classrooms+' '+ line + ' '
                            continue #loop until all classrooms are added
                    
                    parts = re.split(r'(\s\d\s)', line) #split line based on semester (whitespace1whitespace)
                    
                    classrooms=classrooms+parts[0]
                    semester=parts[1]
                    line=parts[2]
                    
                    lessonFound=True
                    
                    print("classrooms: " + classrooms)
                    
            line=line.strip()
            
            exam_class_split = r"([Α-Ω]\s-\s[Α-Ω])"  # Matches "Α - Ω" pattern
                
            #searches for "Α - Ω" in rest of line
            if not re.search(exam_class_split, line): #"Α - Ω" not found so whole line is lesson_name
                lesson_name= lesson_name + ' ' + line #appends to lesson name in case of many iterations
                extendedDescriptionFound=True
                continue #read next line, loops until "Α - Ω" pattern is found
            
            parts = re.split(exam_class_split, line) #splits line based on "Α - Ω" pattern
            lesson_name= (lesson_name + ' ' + parts[0])
            
            exam_class=parts[1].strip()
            lesson_code= parts[2]

            lesson_name=delete_unnecessary_text(lesson_name, lesson_code)
            print("lesson name: "+lesson_name)
            print("exam class: "+exam_class)
            
            lesson_code= re.sub(r'\s', '', parts[2]) #removes all whitespace from lesson_code
            print("lesson code: "+ lesson_code)

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
            exam_time=''
            semester=''
            lesson_code=''
                    

    return exam_data

pdf_file_path = 'dates.pdf'
exam_data = extract_exam_data(pdf_file_path)
df = pd.DataFrame(exam_data)
df.to_excel('exam_data.xlsx', index=False)

create_exam_calendar(exam_data)