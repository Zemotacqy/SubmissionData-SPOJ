import sys, webbrowser, requests, bs4, xlsxwriter

listOfUsers = {}

for i in range(1, len(sys.argv)):
    name = sys.argv[i]
    #print(name)

    res = requests.get('http://www.spoj.com/users/' + name + '/')

    try:
        res.raise_for_status()
    except Exception as exc:
        print('There was a problem')

    soup = bs4.BeautifulSoup(res.text, 'html.parser')
    
    questionlist = soup.select('#user-profile-tables a')
    listOfQuestions = {}
    
    for question in questionlist:
        if(len(question.getText())>0):
            #listOfQuestions.append(question.getText())
            #print(question.getText())

            questionDataRequest = requests.get('http://www.spoj.com/status/' + question.getText() + ',' + name)
            questionSoup = bs4.BeautifulSoup(questionDataRequest.text, 'html.parser')

            listOfSubmission = questionSoup.select('.problems tbody tr')
            submission = []
            for i in range(len(listOfSubmission)):
                keyValues = { 'submission_id', 'submissionTimestamp', 'submissionStatus', 'solutionTime', 'solutionLanguage' }
                submissionObject = dict.fromkeys(keyValues)
                submission_id = questionSoup.select('.problems tbody tr td')[0].getText().strip().strip('\n').strip('\t')
                #print(submission_id)
                submissionObject['submission_id'] = submission_id

                submissionTimestamp = questionSoup.select('.problems tbody tr .status_sm')[i].getText().strip().strip('\n').strip('\t')
                #print(submissionTimestamp)
                submissionObject['submissionTimestamp'] = submissionTimestamp

                
                submissionStatus = questionSoup.select('.problems tbody tr .statusres')[i].getText().strip().strip('\n').strip('\t')
                #print(submissionStatus)
                submissionObject['submissionStatus'] = submissionStatus

                solutionTime = questionSoup.select('.problems tbody tr .stime')[i].getText().strip().strip('\n').strip('\t')
                #print(solutionTime)
                submissionObject['solutionTime'] = solutionTime

                solutionLanguage = questionSoup.select('.problems tbody tr .slang')[i].getText().strip().strip('\n').strip('\t')
                #print(solutionLanguage)
                submissionObject['solutionLanguage'] = solutionLanguage
                submission.append(submissionObject)
                
            #print(submissionObject)
            listOfQuestions[str(question.getText())] = submission
            #print(listOfQuestions)
        
    workbook = xlsxwriter.Workbook('./worksheets/'+name+'.xlsx')
    worksheet = workbook.add_worksheet()
    boldFormat = workbook.add_format({ 'bold' : True })
    questionColorFormat = workbook.add_format({'bold':True, 'font_color':'blue', 'bg_color':'gray'})
    acceptedColorFormat = workbook.add_format({'bg_color':'green'})
    wrongColorFormat = workbook.add_format({'bg_color':'orange'})
    row = 0
    col = 0
    worksheet.write(row, col+2, name, boldFormat)
    row = row + 2
    
    for key in listOfQuestions:
        col=0
        worksheet.write(row, col+2, key, questionColorFormat)
        row = row+1
        worksheet.write(row, col, 'Submission_ID', boldFormat)
        col = col+1
        worksheet.write(row, col, 'TimeStamp', boldFormat)
        col = col+1
        worksheet.write(row, col, 'Status', boldFormat)
        col = col+1
        worksheet.write(row, col, 'Time', boldFormat)
        col = col+1
        worksheet.write(row, col, 'Language', boldFormat)
        col = col+1
        row = row+1
        for submissions in listOfQuestions[key]:
            bgColorFormat = workbook.add_format({})
            if(submissions['submissionStatus'] == 'accepted'):
                bgColorFormat = acceptedColorFormat
                
            if(submissions['submissionStatus'] == 'wrong answer'):
                bgColorFormat = wrongColorFormat
            col=0
            worksheet.write(row, col, submissions['submission_id'], bgColorFormat)
            col = col+1
            worksheet.write(row, col, submissions['submissionTimestamp'], bgColorFormat)
            col = col+1
            worksheet.write(row, col, submissions['submissionStatus'], bgColorFormat)
            col = col+1
            worksheet.write(row, col, submissions['solutionTime'], bgColorFormat)
            col = col+1
            worksheet.write(row, col, submissions['solutionLanguage'], bgColorFormat)
            col = col+1
            row = row + 1
        row = row + 2

    workbook.close()
    print(name+'.xlsx Created Successfully!!!!')
    listOfUsers[str(name)] = listOfQuestions
    
#print(listOfUsers)
