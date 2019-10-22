from django.shortcuts import render
from django.conf import settings
from django.core.files.storage import FileSystemStorage
from django.http import HttpResponse, Http404
from wsgiref.util import FileWrapper
import mimetypes
import os
from django.utils.encoding import smart_str
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
import xlwt
from xlrd import open_workbook

# from . import mlcode

def run_machine_learning(file,dataset):
	# Importing the libraries


	# Excel Sheet parsing
	book=open_workbook(file)
	sheet=book.sheet_by_index(0)
	wb=xlwt.Workbook()
	ws=wb.add_sheet("results")
	data=[]
	for row in range(1,sheet.nrows):
		row_data=[]
		for col in range(0,sheet.ncols):
			if(isinstance(sheet.cell(row,col).value,float)):
				row_data.append(int(sheet.cell(row,col).value))
				ws.write(row,col,int(sheet.cell(row,col).value))
			else:
				row_data.append(sheet.cell(row,col).value)
				ws.write(row,col,sheet.cell(row,col).value)
		data.append(row_data)
	col_names=['Age', 'Attrition', 'Travel', 'DR', 'Dept', 'DistanceFromHome', 'Ed',
	       'EducationField', 'EnvSatisfaction', 'Gender', 'HourlyRate',
	       'Involvement', 'JobLevel', 'Role', 'Satisfaction', 'Marital', 'Income',
	       'MonthlyRate', 'NumCompaniesWorked', 'OverTime', 'PercentSalaryHike',
	       'RelationshipSatisfaction', 'StockOptionLevel', 'TotalWorkingYears',
	       'TrainingTimesLastYear', 'WorkLifeBalance', 'YearsAtCompany',
	       'YearsInCurrentRole', 'YearsSinceLastPromotion', 'YearsWithCurrManager']
	for col,col_val in enumerate(col_names):
		ws.write(0,col,col_val)
	data_frame= pd.DataFrame(data,columns=col_names)
	data_frametemp = data_frame

	# # Importing the dataset
	# dataset = pd.read_csv('employee_data_modified.csv')

	# Data Preprocessing
	dataset.corr()

	col=dataset.columns
	num_col=dataset._get_numeric_data().columns
	list(set(col)-set(num_col))

	var = ['Dept','Role','EducationField','Attrition','Marital','OverTime', 'Gender', 'Travel']

	from sklearn.preprocessing import LabelEncoder
	labelEncoder = LabelEncoder()
	for i in var:
	    dataset[i]=labelEncoder.fit_transform(dataset[i])

	dataset.columns
	X = dataset.loc[: , ['Age', 'Attrition', 'Travel', 'DR', 'Dept', 'DistanceFromHome', 'Ed',
	       'EducationField', 'EnvSatisfaction', 'Gender', 'HourlyRate',
	       'Involvement', 'JobLevel', 'Role', 'Satisfaction', 'Marital', 'Income',
	       'MonthlyRate', 'NumCompaniesWorked', 'OverTime', 'PercentSalaryHike',
	       'RelationshipSatisfaction', 'StockOptionLevel', 'TotalWorkingYears',
	       'TrainingTimesLastYear', 'WorkLifeBalance', 'YearsAtCompany',
	       'YearsInCurrentRole', 'YearsSinceLastPromotion', 'YearsWithCurrManager']]

	y = dataset.loc[: , ['PerformanceRating']]

	# Train Test Split 
	from sklearn.model_selection import train_test_split
	X_train, X_test, y_train, y_test = train_test_split(X, y, test_size = 0.3, random_state = 42, stratify=dataset.PerformanceRating)


	# RandomForestClassifier
	from sklearn.ensemble import RandomForestClassifier
	rf = RandomForestClassifier(random_state=0)
	rf.fit(X_train,y_train)

	#Preprocessing the data given through website
	var = ['Dept','Role','EducationField','Attrition','Marital','OverTime', 'Gender', 'Travel']

	from sklearn.preprocessing import LabelEncoder
	labelEncoder = LabelEncoder()
	for i in var:
	    data_frametemp[i]=labelEncoder.fit_transform(data_frametemp[i])
	y_pred=rf.predict(data_frametemp)

	# reconstructing the sheet with the new column
	ws.write(0,sheet.ncols,"PerformanceRating")
	for row in range(1,len(data)+1):
		ws.write(row,sheet.ncols,int(y_pred[row-1]))
	wb.save('media/result.xlsx')
	
	# # # Scores
	# # from sklearn.metrics import confusion_matrix
	# # cm = confusion_matrix(y_test, y_pred)

	# # from sklearn.metrics import accuracy_score
	# # print(accuracy_score(y_test,y_pred))


def homepage(request):
	if request.method=="GET":
		return render(request,'mainapp/index.html')
	if request.method == 'POST' and request.FILES['myfile']:
	    myfile = request.FILES['myfile']
	    fs = FileSystemStorage()
	    filename = fs.save(myfile.name, myfile)
	    uploaded_file_url = fs.url(filename)

	    file_path = settings.MEDIA_ROOT +'/'+ filename
	    dataset=pd.read_csv('mainapp/employee_data_modified.csv')
	    run_machine_learning(file_path,dataset)
	    file_path = settings.MEDIA_ROOT +'/'+ "result.xlsx"
	    file_wrapper = FileWrapper(open(file_path,'rb'))
	    file_mimetype = mimetypes.guess_type(file_path)
	    response = HttpResponse(file_wrapper, content_type=file_mimetype )
	    response['X-Sendfile'] = file_path
	    response['Content-Length'] = os.stat(file_path).st_size
	    response['Content-Disposition'] = 'attachment; filename=%s' % smart_str("result.xlsx") 

	    return response

	    # return render(request, 'mainapp/index.html', {
	    #     'uploaded_file_url': uploaded_file_url
	    # })
