# --------------------------Project Description-------------------#
"""
 About : Resume Shortlist system is used for fetching best resumes out of bulk aon the bases of JD.
 Technology : NLP / Regex / Pandas clening
 Version : python 3.7
 Organization : Proven Consult
 Auther : Shoeb Ahmad ( Prof < sahmad@provenconsult.com > , Outside < er.shoaib10@gmail.com > )

"""
# --------------------------How to Use----------------------------#
"""
Step 1: Run script it will create respected folder automatically once done ,then code will quit 
Step 2: Placed all three (education.txt / Relavent Experiance.txt / Skills.txt) corpus file in corpus folder
Step 3: Placed All resumes in Input-Resumes folder
step 4: Placed JD in JD-And-Output folder

"""
# ----------------------------Packages Used-----------------------#

# Try to fetch required package if any package is missing code will exit

try:
    import spacy
    import os ,time
    import shutil
    import re
    import calendar
    from datetime import datetime
    import pandas as pd
    from StyleFrame import StyleFrame, Styler, utils
    import numpy as np
    from spacy_lookup import Entity
    from textract import process
    from operator import add
    import subprocess
    import warnings
    warnings.filterwarnings("ignore")
except ModuleNotFoundError or Exception as e:
    print(e)
    exit()

# -----------------CODE MAIN BODY__________________________________#

# Creating directories if not in path


execution_path = os.getcwd()
def create_direcory():
    try:
        def Check_Path(Path):
            try:
                if not os.path.exists(Path):
                    os.mkdir(Path)
                    print(Path + " " + 'is successfully created')
                else:
                    pass
            except Exception as ed:
                print(ed)

        Check_Path(execution_path + "//" + "Input-Resumes")
        Check_Path(execution_path + "//" + "JD-And-Output")
        Check_Path(execution_path + "//" + "corpus")
        time.sleep(2)
        Check_Path(execution_path + "//" + "JD-And-Output" + '//' + 'Selected Resumes')

        # Tool Input and Output Directory

        INPATH = execution_path + "//" + "Input-Resumes"
        OUTPATH = execution_path + "//" + "JD-And-Output"
        corpus =  execution_path + "//" + "corpus"
        selected_Resume = execution_path + "//" + "JD-And-Output" + '//' + 'Selected Resumes'
        return INPATH ,OUTPATH ,corpus ,selected_Resume
    except (PermissionError,IsADirectoryError) as e:
        print(e)

INPATH , OUTPATH ,corpus, selected_Resume = create_direcory()

# Before Processing below module check is their any file is present in directory, if not then code will exit()

def check_for_files_in_ResumeInput_Folder():
    try:
        if not os.listdir(INPATH) or not os.listdir(corpus):
            print('No File Found For Processing')
            exit()
        else:
            pass
    except (NotADirectoryError,FileNotFoundError,Exception) as e:
        print(e)
    return None
check_for_files_in_ResumeInput_Folder()
Names=[]

def clear_selected_resume_folder():
    folder = selected_Resume
    for the_file in os.listdir(folder):
        file_path = os.path.join(folder, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            #elif os.path.isdir(file_path): shutil.rmtree(file_path)
        except Exception as e:
            print(e)
    return None
clear_selected_resume_folder()
# Doc to Docx convertion

def doc_to_docx(Args):
    try:
        for doc in os.listdir(Args):
            if doc.lower().endswith('doc'):
                subprocess.call(['soffice', '--headless', '--convert-to', 'docx', Args + '/' + doc])
        print('Doc to Docx conversion Done')
        time.sleep(2)
        for move in os.listdir(execution_path):
            if move.lower().endswith('docx'):
                shutil.move(execution_path+'/'+move,INPATH)
        for Del in os.listdir(Args):
            if Del.lower().endswith('doc'):
                os.unlink(Args+'/'+Del)
    except FileNotFoundError as e:
        print(e)

doc_to_docx(INPATH)
# Module for read resume data and append in list for further use
def read_resumes(Args):
    data=[]
    for resumes in os.listdir(Args):
        if resumes.lower().endswith('pdf') or resumes.lower().endswith('docx') or resumes.lower().endswith('dox'):
            if Args == INPATH:
                Names.append(resumes)
            resume=process(Args+'/'+resumes)
            read_resume=(resume.decode("utf-8").replace(',',' ').replace('/',' ').replace('\n',' ').replace('Data Science','Data-Science').replace('Data Scientist','Data-Scientist')
                         .replace('data scientist','data-scientist').replace('data science','data-science').lower())
            resume_content=(' '.join(read_resume.split()))
            data.append(resume_content)
    return data

# Module for reading skills from csv file for trained model as per NER

def read_skill_Train_Data(File):
    file = open(File, "r")

    return file.readline().lower()

# Module for tokenizing skills before get it trained for entity

def tokenization(Args):
    tokenization=Args.replace(', ',',').split(',')
    return (tokenization)

# Module for entity training through en model by making all pipeline disable

def tarined_entity_for_skills(Args):
    Skill_Entity_Model=spacy.load('en',disable = ['ner', 'tagger', 'parser', 'textcat'])
    Skill_entity = Entity(keywords_list=Args, label='Skill')
    Skill_Entity_Model.add_pipe(Skill_entity, last=True)
    return Skill_Entity_Model

# Module for extracting skills from resumes

def Extract_Using_NLP(Args,Skill):
    Exracted_Skills =[]
    File=read_skill_Train_Data(Skill)
    Tokenize_data=tokenization(File)
    #print(Tokenize_data)
    Skill_Entity_Model = tarined_entity_for_skills(Tokenize_data)
    Skill_data = Skill_Entity_Model(Args)
    for skills in Skill_data.ents:
        #print(skills)
        Exracted_Skills.append(skills.text)
    return list(set(Exracted_Skills))

# Module for extracting phone no from resumes

def Extract_Mobile_Number(inputString):
    Mobile_Numbers = []
    r = re.compile(r'(\d{3}[-\.\s]??\d{3}[-\.\s]??\d{4}|\(\d{3}\)\s*\d{3}[-\.\s]??\d{4}|\d{3}[-\.\s]??\d{4})')
    phone_numbers = r.findall(inputString)
    get_number=([re.sub(r'\D', '', number) for number in phone_numbers])
    for number in get_number:
        if len(number)>=8:
            Mobile_Numbers.append(number)
    return Mobile_Numbers

# module for extracting email from resumes

def extract_email(email):
    #print(email)
    Email_ids=[]
    email = re.findall("([^@|\s]+@[^@]+\.[^@|\s]+)", email)
    try:
        if email:

                email= email[0].split()[0].strip(';')
                Email_ids.append(email)
                #print(email)
                return Email_ids
        elif not email:
                Email_ids.append('No Email Found in Resume')
                return Email_ids
    except Exception as e:
        print(e)


def Immediate_Joining(data):
    Joining_Status=[]
    try:
        Immediate_Corpus=['Join Immediate','Immediate Join','Immediate Available','Available Immediate']
        for check_Immediate_Joine in Immediate_Corpus:
            if check_Immediate_Joine.lower() in data.lower():
                Joining_Status.append('Immediate Available')
                break
            else:
                Joining_Status.append('Not Immediate Available')
        return Joining_Status[-1]
    except Exception as e:
        print(e)

def extract_experence(Args,Args1):
    exp = []
    data = Args.replace('-', ' ').replace('+',' ').replace('.','').split()


    try:
        if data != [] or data != None:
            for check in range(len(data)):
                #print(check)
                if data[check] == 'year' or data[check] == 'years' or data[check] == 'yrs':
                    if data[check - 1].isdigit():
                        # condition check becoz many time it will fetch extra no so this check help to round off
                        if len(data[check - 1])>2:
                            data[check - 1]=data[check - 1][:2]
                        else:
                            pass
                        exp.append(str(int(data[check - 1]).__round__()))
                    elif data[check - 2].isdigit():
                        # condition check becoz many time it will fetch extra no so this check help to round off
                        if len(data[check - 2])>2:
                            data[check - 2]=data[check - 2][:2]
                        else:
                            pass
                        exp.append(str(int(data[check - 2]).__round__()))
                    elif data[check + 1].isdigit():
                        # condition check becoz many time it will fetch extra no so this check help to round off
                        if len(data[check + 1])>2:
                            data[check + 1]=data[check + 1][:2]
                        else:
                            pass
                        exp.append(data[check + 1])
                    elif data[check + 2].isdigit():
                        # condition check becoz many time it will fetch extra no so this check help to round off
                        if len(data[check + 1])>2:
                            data[check + 1]=data[check + 1][:2]
                        else:
                            pass
                        exp.append(data[check + 2])
                    else:
                        exp.append(str(0))


        else:
            pass
        return Args1(exp, default=str(0))

    except ValueError as e:
        print(e,'190')
        #exp.append(str('NA'))
        pass

# Relevent experiance parser
# Extracting search word for relavent experiance
def check():
    CL=[]
    try:
        for name in os.listdir(OUTPATH):
            if not name.endswith('xlsx'):
                namecheck=(name.replace('-','').replace("'", '').replace(',', '').replace('(', '').replace( ')', '')
                .replace('[', '').replace(']', '').replace('+', '').replace('  ',' ').replace('.',' ').lower().split())
                for NC in namecheck[:-1]:
                    CL.append(NC)
        return CL
    except Exception as e:
        print(e)

# Function used to fetch relevant experiance from resumes

def Relevent_Exp_parser(data):
    test = []
    temparr = []
    relaventExp = []
    check1array = []

    # below we extract skills from jd in order to use it for fetching correct date of relevant experience
    try:
        for JD_data in JD:
            Skills_in_JD = Extract_Using_NLP(JD_data, corpus + '/' + 'Skills')
            test.append(Skills_in_JD)
    except Exception as e:
        print(e)

    try:
        def hasNumbers(inputString):     # func to find number in string
            return bool(re.search(r'\d', inputString))
        # main array in which all extraction are stored
        # below func check that the word in corpur are present in resume or not
        Relavend_Exp = Extract_Using_NLP(data.lower(), corpus + '/' + 'Relavent Experiance')
        if Relavend_Exp != []:
            # data is in string format are re structured to list for perform below operation
            ExtractedData = data.replace(':', '').replace('(', ' ').replace(')', ' ').replace('+', '').replace('.', '').replace('-to',' - to ').replace('-till',' - till ')\
                .replace('-present',' - present ').replace('+',' ').replace('-present',' - present ').split()
            # print(ExtractedData)
            # below is decision based string if assig value is still 0 then sec conf only iterate if 1 then first .
            CondCheck=0
            #print(test[0])
            for x in range(len(ExtractedData[:120])):

                if (test[0]).__contains__(ExtractedData[:120][x]):
                    #print(ExtractedData[:150][x],'pop')
                    #print(ExtractedData[:150][x-7:x+10])
                    #first condition for fetching relevant exp from resumes where person mention exp in top section that why
                    #loop iterate up to range 150 only
                    # below is the first condition and the loop will iterate only when find Rex under starting 150 word then sec will not no A become 1
                    for n in  (ExtractedData[0:120][x-10:x+5]): # if yes then check 5 sentence before and after from the word
                        # below condition chech in year or years word find in ExtractedData[0:120][x-7:x+10] then go forward

                        if n.__contains__('year') or n.__contains__('years') or n.__contains__('yrs'):
                            # below code have 4 condition check , for fetching exact no of relevant experiance
                            for x1 in range(len(ExtractedData[0:120][x-7:x+10])):

                                if (ExtractedData[0:120][x-7:x+10][x1]=='year' or  ExtractedData[0:120][x-7:x+10][x1]=='yrs' or ExtractedData[0:120][x-7:x+10][x1]=='years') and hasNumbers(ExtractedData[0:120][x-7:x+10][x1-1])==True:
                                    #print(ExtractedData[0:120][x - 7:x + 10][x1 - 1],'poip')
                                    # below condition will check if any elem with len 2 and greater then 20 is present so rearrange so that it will write in correct format
                                    if len(ExtractedData[0:120][x-7:x+10][x1-1])==2 and int(ExtractedData[0:120][x-7:x+10][x1-1])>20:
                                        CondCheck=1
                                        if ExtractedData[0:120][x-7:x+10][x1-1][0]+'.'+ExtractedData[0:120][x-7:x+10][x1-1][1] not in relaventExp:
                                            relaventExp.append(ExtractedData[0:120][x-7:x+10][x1-1][0]+'.'+ExtractedData[0:120][x-7:x+10][x1-1][1])
                                    # below condition will check if any elem with len 2 and less then 20 is present so rearrange so that it will write in correct format
                                    elif len(ExtractedData[0:120][x-7:x+10][x1-1])==2 and int(ExtractedData[0:120][x-7:x+10][x1-1])<20:
                                        CondCheck = 1
                                        if ExtractedData[0:120][x-7:x+10][x1-1] not in relaventExp:
                                            relaventExp.append(ExtractedData[0:120][x-7:x+10][x1-1])
                                    elif len(ExtractedData[0:120][x - 7:x + 10][x1 - 1]) == 1:
                                        CondCheck = 1
                                        #print(ExtractedData[0:120][x - 7:x + 10][x1 - 1])
                                        if ExtractedData[0:120][x - 7:x + 10][x1 - 1] not in relaventExp:
                                            relaventExp.append(ExtractedData[0:120][x - 7:x + 10][x1 - 1])

                                elif ExtractedData[0:120][x-7:x+10][x1]=='year' or (ExtractedData[0:120][x-7:x+10][x1]=='years' and hasNumbers(ExtractedData[0:120][x-7:x+10][x1-2])==True):

                                    # below condition will check if any elem with len 2  and greater then 20 is present so rearrange so that it will write in correct format
                                    if len(ExtractedData[0:120][x-7:x+10][x1-2])==2 and int(ExtractedData[0:120][x-7:x+10][x1-2])>20:
                                        CondCheck = 1
                                        if ExtractedData[0:120][x-7:x+10][x1-2][0]+'.'+ExtractedData[0:120][x-7:x+10][x1-2][1] not in relaventExp:
                                            relaventExp.append(ExtractedData[0:120][x-7:x+10][x1-2][0]+'.'+ExtractedData[0:120][x-7:x+10][x1-2][1])
                                    # below condition will check if any elem with len 2 and less then 20 is present so rearrange so that it will write in correct format
                                    elif len(ExtractedData[0:120][x-7:x+10][x1-2])==2 and int(ExtractedData[0:120][x-7:x+10][x1-2])<20:
                                        CondCheck = 1
                                        if ExtractedData[0:120][x - 7:x + 5][x1 - 2] not in relaventExp:
                                            relaventExp.append(ExtractedData[0:120][x-7:x+10][x1-2])
                                    elif len(ExtractedData[0:120][x - 5:x + 5][x1 - 2]) == 1:
                                        CondCheck = 1
                                        if ExtractedData[0:120][x - 5:x + 5][x1 - 2] not in relaventExp:
                                            relaventExp.append(ExtractedData[0:120][x - 5:x + 5][x1 - 2])
                                else:
                                    pass
                            break
            #print(CondCheck)
            # condition 2 will checkin whole resume content and search for period base exp data

                # below loop will check for till and present word in data if getting same then append only some reagion of +-5 in check1array=[] for further check
            #print(CondCheck)
            A = 0
            if CondCheck == 0:
                for cond2 in range(len(ExtractedData)):
                    Month_Copus1 = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'july', 'aug', 'sep', 'oct', 'nov', 'dec',
                                    'january',
                                    'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september',
                                    'october', 'december']

                    if ExtractedData[cond2] == ('till') or ExtractedData[cond2] == ('present') :
                        if A ==0:
                            if ( ExtractedData[cond2 - 1].isnumeric() and len(ExtractedData[cond2 - 1])<=4) or ( ExtractedData[cond2 - 2].isnumeric() and len(ExtractedData[cond2 - 2])<=4):
                                    A = 1
                                    #print(ExtractedData[cond2-10:cond2+150])
                                    for i2 in  ExtractedData[cond2-10:cond2+150]:
                                        if test[0]!=[] and test[0].__contains__(i2):
                                            check1= ExtractedData[cond2-3:cond2+1]

                                            if check1  not in check1array:

                                                #print(check1)
                                                check1array.append(check1)
                                            else:
                                                break
                    elif ExtractedData[cond2] == ('–') or ExtractedData[cond2] == ('to') or ExtractedData[cond2] == ('-'):
                        for month in Month_Copus1:
                            if ExtractedData[cond2 - 1].isnumeric() and (
                                   ExtractedData[cond2 + 1].__contains__(month) or ExtractedData[cond2 + 2].__contains__(month) or ExtractedData[cond2 + 3].__contains__(month)):
                            #if ExtractedData[cond2-1].isnumeric()  and (Month_Copus1.__contains__(ExtractedData[cond2+1]) or Month_Copus1.__contains__(ExtractedData[cond2+2]) or Month_Copus1.__contains__(ExtractedData[cond2+3])):
                                for i in (ExtractedData[cond2-10:cond2+150]):
                                    if test[0]!=[] and test[0].__contains__(i):
                                        # print(test[0])
                                        #print((ExtractedData[cond2 - 2:cond2 + 4]))
                                        temparr.append(ExtractedData[cond2-2:cond2+3])
                                        break

                if temparr != None and len(temparr)>1:
                    def temptest1():
                        for arr in temparr[0]:
                            if arr == '–' or arr == '–' or arr == '-' or arr == 'to':
                                part1=(temparr[0][temparr[0].index(arr):])
                                #print(part1)
                                return part1
                    for arr in temparr[-1]:
                        part1=temptest1()
                        if part1!=None and (arr == '–' or arr == '–' or arr == 'to'):
                            part2 = (temparr[-1][:temparr[-1].index(arr)+1])
                            #print(part2)
                            check1array.append(part2+['<-to->']+part1)
                            #print(part1+['to']+part2,'part')
                        else:
                            pass
                elif len(temparr)==1:
                    #print(temparr[-1])
                    check1array[-1].append(temparr+['!'])
                #print(check1array,'opo')
                if check1array != None and len(check1array)>=2:
                    for get_num in check1array:
                        #print(get_num,'pop')
                        # below condition check if number present in list and not having both keywords till and present
                        if hasNumbers(str(get_num)) == True  and (get_num.__contains__('till') or get_num.__contains__('present')):
                            if get_num.__contains__('till'):
                                if len(temparr) !=0:
                                    filterdata = check1array[-1]+["AND"]+get_num[:get_num.index('till')+2]
                                    #print(filterdata,'filter')

                                    relaventExp.append(filterdata)

                                else:
                                    filterdata = get_num[:get_num.index('till') + 2]
                                    relaventExp.append(filterdata)

                            elif get_num.__contains__('present'):
                                # print(get_num)
                                if len(temparr)!=0:
                                    filterdata = check1array[-1]+["AND"]+get_num[:get_num.index('present')+1]
                                    #print(filterdata,'llllllllll')
                                    relaventExp.append(filterdata)
                                else:
                                    filterdata = get_num[:get_num.index('present') + 2]
                                    relaventExp.append(filterdata)
                elif check1array != None:
                    #print('pp')
                    relaventExp.append(check1array)
                    '''elif get_num.__contains__('<-to->'):
                        #print(get_num,'popopop')
                        relaventExp.append(get_num)
                    elif get_num.__contains__('!'):
                        relaventExp.append(get_num[0])'''


        return relaventExp
    except Exception as e:
        print(e,'279')

# Module to check stability of candidate

def stability(data):
    Temp_list=[]
    Stability=[]
    # First Sprit data and make it formated
    ExtractedData = data.replace(':', '').replace('(', '').replace(')', '').replace('+', '').replace('-', ' ').replace('.', '').split()
    # below condition will check if data conatin duration word and with in 10 word from duration if contain till or period then make it separate
    if ExtractedData.__contains__('duration') and \
            (ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10].__contains__('till') or
             ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10].__contains__('present')):
        # Below if/else will check if contail till or else if contain present then append 3 words perior to that in temp list for furter operations
        if (ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10]).__contains__('till'):
            Temp_list.append(ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10]
                  [ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10].index('till')-3:
            ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10].index('till')+1])
        elif (ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10]).__contains__('present'):
            Temp_list.append(ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10]
                  [ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10].index('present')-3:
            ExtractedData[ExtractedData.index('duration'):ExtractedData.index('duration')+10].index('present')+1])

    # now trying to fetch appropriate data like start month , end month and year from data stored in Temp_List

    #---------- Corpus for required operation------------------#````
    # below logic will convert nomenclature of month into number
    abbr_to_num = {name: num for num, name in enumerate(calendar.month_abbr or calendar.month_name) if num}
    # dict for month to month_NMC
    Month_dict = {"january": "jan", "february": "feb", "march": "mar", "april": "apr", "may": "may", "june": "jun",
                  "july": "jul", "august": "aug", "september": "sep", "october": "oct", "november": "nov",
                  "december": "dec"}
    # below variable and list are define to perform further logical operation
    current = datetime.now()

    present = current.year

    Month_Copus1 = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'july', 'aug', 'sep', 'oct', 'nov', 'dec', 'january',
                    'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'december']

    def Date_sub(Args):
        try:
            now = datetime.now()
            todaysdate = (now.strftime("%m/%y"))
            d0 = datetime.strptime(todaysdate, "%m/%y")
            d1 = datetime.strptime(Args, "%m/%y")
            delta = d0 - d1
            output=(int(delta.days / 30))
            return output
        except Exception as e:
            print(e)

    for get in Temp_list:
            #print(get)
            if get.__contains__('till') and (get.__contains__('–') or get.__contains__('to')):
                year = (get[get.index('till') - 2])
                if len(get[get.index('till')-3])>3 and Month_Copus1.__contains__(get[get.index('till')-3]) and len(year)==4:
                    Month=(get[get.index('till')-3])
                    year = (get[get.index('till') - 2])
                    Month_dict = {"january": "jan", "february": "feb", "march": "mar", "april": "apr", "may": "may",
                                  "june": "jun",
                                  "july": "jul", "august": "aug", "september": "sep", "october": "oct",
                                  "november": "nov",
                                  "december": "dec"}
                    date_mon=(str(abbr_to_num[Month_dict[Month].title()])+'/'+year[2:])
                    Month_in_days=Date_sub(date_mon)
                    if Month_in_days >10:
                       Stability.append('No Stability Issue')
                    else:
                        Stability.append('Alert last employee less then'+" "+str(Month_in_days)+" "+"month")
                elif Month_Copus1.__contains__(get[get.indeDatax('till')-3]) and len(year)==4:
                    Month = (get[get.index('till') - 3])
                    year = (get[get.index('till') - 2])
                    date_mon = (str(abbr_to_num[Month.title()]) + '/' + year[2:])
                    Month_in_days = Date_sub(date_mon)
                    if Month_in_days > 10:
                        Stability.append('No Stability Issue')
                    else:
                        Stability.append('Alert last employee less then' + " " + str(Month_in_days) + " " + "month")

            elif get.__contains__('till') and not (get.__contains__('–') or get.__contains__('to')):
                year = (get[get.index('till') - 1])
                if len(get[get.index('till') - 2]) > 3 and Month_Copus1.__contains__(
                        get[get.index('till') - 2]) and len(year) == 4:
                    Month = (get[get.index('till') - 2])
                    year = (get[get.index('till') - 1])
                    Month_dict = {"january": "jan", "february": "feb", "march": "mar", "april": "apr", "may": "may",
                                  "june": "jun",
                                  "july": "jul", "august": "aug", "september": "sep", "october": "oct",
                                  "november": "nov",
                                  "december": "dec"}
                    date_mon = (str(abbr_to_num[Month_dict[Month].title()]) + '/' + year[2:])
                    Month_in_days = Date_sub(date_mon)
                    if Month_in_days > 10:
                        Stability.append('No Stability Issue')
                    else:
                        Stability.append('Alert last employee less then' + " " + str(Month_in_days) + " " + "month")
                elif Month_Copus1.__contains__(get[get.index('till') - 2]) and len(year) == 4:
                    Month = (get[get.index('till') - 2])
                    year = (get[get.index('till') - 1])
                    date_mon = (str(abbr_to_num[Month.title()]) + '/' + year[2:])
                    Month_in_days = Date_sub(date_mon)
                    if Month_in_days > 10:
                        Stability.append('No Stability Issue')
                    else:
                        Stability.append('Alert last employee less then' + " " + str(Month_in_days) + " " + "month")

            elif  get.__contains__('present') and (get.__contains__('–') or get.__contains__('to')):
                year = (get[get.index('present') - 2])
                if len(get[get.index('present') - 3]) > 3 and Month_Copus1.__contains__(get[get.index('present') - 3]) and len(year) == 4:
                    Month = (get[get.index('present') - 3])
                    year = (get[get.index('present') - 2])
                    Month_dict = {"january": "jan", "february": "feb", "march": "mar", "april": "apr", "may": "may",
                                  "june": "jun",
                                  "july": "jul", "august": "aug", "september": "sep", "october": "oct",
                                  "november": "nov",
                                  "december": "dec"}
                    date_mon = (str(abbr_to_num[Month_dict[Month].title()]) + '/' + year[2:])
                    Month_in_days = Date_sub(date_mon)
                    if Month_in_days > 10:
                        Stability.append('No Stability Issue')
                    else:
                        Stability.append('Alert last employee less then' + " " + str(Month_in_days) + " " + "month")
                elif Month_Copus1.__contains__(get[get.index('present')-3]) and len(year)==4:
                    Month = (get[get.index('present') - 3])
                    year = (get[get.index('present') - 2])
                    date_mon = (str(abbr_to_num[Month.title()]) + '/' + year[2:])
                    Month_in_days = Date_sub(date_mon)
                    if Month_in_days > 10:
                        Stability.append('No Stability Issue')
                    else:
                        Stability.append('Alert last employee less then' + " " + str(Month_in_days) + " " + "month")

            elif  get.__contains__('present') and not (get.__contains__('–') or get.__contains__('to')):
                year = (get[get.index('present') - 1])
                if len(get[get.index('present') - 2]) > 3 and Month_Copus1.__contains__(get[get.index('present') - 2]) and len(year) == 4:
                    Month = (get[get.index('present') - 2])
                    year = (get[get.index('present') - 1])
                    Month_dict = {"january": "jan", "february": "feb", "march": "mar", "april": "apr", "may": "may",
                                  "june": "jun",
                                  "july": "jul", "august": "aug", "september": "sep", "october": "oct",
                                  "november": "nov",
                                  "december": "dec"}
                    date_mon = (str(abbr_to_num[Month_dict[Month].title()]) + '/' + year[2:])
                    Month_in_days = Date_sub(date_mon)
                    if Month_in_days > 10:
                        Stability.append('No Stability Issue')
                    else:
                        Stability.append('Alert last employee less then' + " " + str(Month_in_days) + " " + "month")
                elif Month_Copus1.__contains__(get[get.index('present')-2]) and len(year)==4:
                    Month = (get[get.index('present') - 2])
                    year = (get[get.index('present') - 1])
                    date_mon = (str(abbr_to_num[Month.title()]) + '/' + year[2:])
                    Month_in_days = Date_sub(date_mon)
                    if Month_in_days > 10:
                        Stability.append('No Stability Issue')
                    else:
                        Stability.append('Alert last employee less then' + " " + str(Month_in_days) + " " + "month")

    return Stability



# call all modules for extracting data from resumes

Resumes=read_resumes(INPATH)
JD=read_resumes(OUTPATH)
Mobile=[]
Email_Address=[]
Immediate_Joining_Status=[]
Resume_Education=[]
Resume_Exp=[]
Resumes_Skills=[]
Relevent_Experiance=[]
Stability_Array=[]

for data in Resumes:
    MobileNo = Extract_Mobile_Number(data)
    Mobile.append((''.join( repr(e) for e in MobileNo ) ))
    Skills_in_Resumes = Extract_Using_NLP(data.lower(),corpus+'/'+'Skills')
    Resumes_Skills.append(( ", ".join( repr(e) for e in Skills_in_Resumes ) ))
    Resume_Education_Detail = Extract_Using_NLP(data.upper(), corpus + '/' + 'education')
    Resume_Education.append(( ", ".join( repr(e) for e in Resume_Education_Detail ) ))
    Email = extract_email(data)
    REx= Relevent_Exp_parser(data)
    if REx == None:
        REx=''
    #print(REx,'len')
    def hasNumbers(inputString):  # func to find number in string
        return bool(re.search(r'\d', inputString))
    if REx != None and len(REx)>2:
       # print(REx)
        if hasNumbers(REx[0])==True and hasNumbers(REx[-1])==True:
            REx=(float(REx[0]))+float(REx[-1])
            REx=str(REx).replace("'",'')
            Relevent_Experiance.append(REx)
    else:
        if REx != None:
            REx=str(REx).replace('[','').replace(']','').replace(',',' ').replace("'",' ')
            Relevent_Experiance.append(REx)
    '''stabilty_check=stability(data)
    Stability_Array.append(str(stabilty_check).replace('[','').replace(']','').replace(',','').replace("'",''))'''
    Email_Address.append(Email[0])
    Resume_Experiance=extract_experence(data,max)
    if Resume_Experiance != None:
        Resume_Exp.append(Resume_Experiance.replace('','.')[1:-1])
    else:
        Resume_Exp.append('No Data')
    status=Immediate_Joining(data)
    Immediate_Joining_Status.append(status)
print('Required Parameters are successfully fetched from resume -------------Done ')
#print(Relevent_Experiance)
JD_Skills=[]
JD_Education=[]
JD_Exp=[]
JD_Unmodified_Skill=[]
for JD_data in JD:
    for i in range(len(Names)):
        Skills_in_JD = Extract_Using_NLP(JD_data, corpus + '/' + 'Skills')
        JD_Unmodified_Skill.append(Skills_in_JD)
        JD_Skills.append(( ", ".join( repr(e) for e in Skills_in_JD ) ))
        JD_Education_Detail = Extract_Using_NLP(JD_data.upper(), corpus + '/' + 'education')
        JD_Education.append(( ", ".join( repr(e) for e in JD_Education_Detail ) ))
        JD_Experiance = extract_experence(JD_data,max)
        JD_Exp.append(JD_Experiance)
print('Required Parameters are successfully fetched from JD -------------Done ')
#print(JD_Skills)
def write_to_excel():
    try:
        df = pd.DataFrame.from_dict({'Resume-Names':Names,'Mobile-No':Mobile,'Email-Address':Email_Address,'Immediate Joining':Immediate_Joining_Status,'Resume(Education)':Resume_Education,
                                     'Resume(Experiance)':Resume_Exp,'Resume(Relevant Experiance)':Relevent_Experiance,'JD(Skills)':JD_Skills,'Resume(Skills)':Resumes_Skills})
        df.to_excel(OUTPATH+'/'+'ShortlistResumes.xlsx', header=True, index=False)
    except Exception as e:
        print(e)
write_to_excel()
print('Mathametical calculation started for scoring--------')
def Skill_JD_Vs_Resumes(Args):
    Exracted_Skills =[]
    trained = ((JD_Unmodified_Skill)[0])
    Skill_Entity_Model = tarined_entity_for_skills(trained)
    Skill_data = Skill_Entity_Model(Args)
    for JDS in Skill_data.ents:
        Exracted_Skills.append(JDS.text)
    return list(set(Exracted_Skills))

def SkillScores():
    Read_Resume=read_resumes(INPATH)
    Compared_Skill=[]
    Resumes_Score=[]
    for skillCheck in Read_Resume:
        SKILLS=Skill_JD_Vs_Resumes(skillCheck)
        Compared_Skill.append(SKILLS)
    for Score in Compared_Skill:
        Resumes_Score.append(len(Score))
    return Resumes_Score

ResumeScores = SkillScores()



def Append_Score_to_Excel(Args,Excel,Column_Name):
    try:
        b=pd.DataFrame(Args,columns=[Column_Name])
        # condition Check one
        to_update = {"Sheet1": b}

        # load existing data
        file_name = OUTPATH+'/'+Excel
        excel_reader = pd.ExcelFile(file_name)

        # write and update
        excel_writer = pd.ExcelWriter(file_name)

        for sheet in excel_reader.sheet_names:
            sheet_df = excel_reader.parse(sheet)
            append_df = to_update.get(sheet)

            if append_df is not None:
                sheet_df = pd.concat([sheet_df, append_df], axis=1)

            sheet_df.to_excel(excel_writer, sheet, index=False)

        excel_writer.save()
    except (FileNotFoundError,Exception) as e:
        print(e)
    return None

def Score_in_Percentage():
    percentage = []
    l=(len(JD_Unmodified_Skill[0]))
    for per in ResumeScores:
        Per_formula=((int(per)*100)/l)
        percentage.append(Per_formula.__round__(1))
    return percentage

Resume_Prediction=Score_in_Percentage()

def ResumeScore():
    Score_list=[]
    for get_per in Resume_Prediction:
        Score=((get_per/100)*6)
        Score_list.append(Score.__round__(1))
    return Score_list

finalScore=ResumeScore()

EXP_CAl=[]
FINAL_SCORE=[]
for i in Resume_Exp:
        try:
            if float(i)>=float(JD_Exp[0]):
                EXP_CAl.append(float(100))
            elif float(i)==0.0:
                EXP_CAl.append(int(0))
            else:
                value=float(JD_Exp[0]) - float(i)
                Per_formula1 = ((float(JD_Exp[0])-value) * 100.0) / float(JD_Exp[0])
                if Per_formula1 >=10:
                    EXP_CAl.append(Per_formula1.__round__())
                else:
                    EXP_CAl.append(0)
        except Exception as e:
            print(e)
for final_score in EXP_CAl:
    Score1 = ((final_score / 100) * 4)
    FINAL_SCORE.append(Score1.__round__(1))
#print(EXP_CAl)
#print(FINAL_SCORE)


FINALSCORE=(list( map(add, FINAL_SCORE, finalScore) ))

Resume_Status=[]
for selection in FINALSCORE:
    if selection <5:
        Resume_Status.append('Not Selected')
    else:
        Resume_Status.append('Selected')
#print(Names)
#print(Resume_Status)
print('Scoring process Successfully-----------Done')
Append_Score_to_Excel(Resume_Prediction,'ShortlistResumes.xlsx','Skill_Percentage')
Append_Score_to_Excel(finalScore,'ShortlistResumes.xlsx','Skill_Score out of 6')
Append_Score_to_Excel(EXP_CAl,'ShortlistResumes.xlsx','Experiance Percentage')
Append_Score_to_Excel(FINAL_SCORE,'ShortlistResumes.xlsx','Experiance Score out of 4')
Append_Score_to_Excel(FINALSCORE,'ShortlistResumes.xlsx','FINAL-SCORE out of 10')
Append_Score_to_Excel(Resume_Status,'ShortlistResumes.xlsx','Resume_Status')

def fill_black_cell():
    # df1 is our output excel i.e ShortlistResumes.xlsx
    # df2 is ShortlistResumes.xlsx after modification
    try:
        def highlight_cells():
            # provide your criteria for highlighting the cells here
            return ['background-color: yellow']
        df1 = pd.read_excel(OUTPATH+'/'+'ShortlistResumes.xlsx')
        df2 = df1.replace(np.nan, 'No Data', regex=True)
        df2.to_excel(OUTPATH+'/'+'ShortlistResumes.xlsx',sheet_name='Resume_Shortlist_Summary')
    except Exception as e:
        print(e)
# this fuction will replace black value of ShortlistResumes.xlsx with no data

fill_black_cell()
print('Output are ready .....enjoy')

# move shortlist remumes to output folder

def Move_selected_resume():
    try:
        resumes = pd.read_excel(OUTPATH+'/'+'ShortlistResumes.xlsx')
        resumes.reset_index()
        selected_resumes=(resumes[resumes["Resume_Status"]=='Selected'])
        Resume_Name=(selected_resumes["Resume-Names"].to_list())
        for get_selection in Resume_Name:
                shutil.copy(INPATH+'/'+get_selection,selected_Resume)
        print('Selected resumes are successfully move to output folder')
    except Exception as e:
        print(e)
    return None
Move_selected_resume()


