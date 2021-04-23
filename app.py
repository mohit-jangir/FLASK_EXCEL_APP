from flask import Flask,render_template,request
import pandas as pd

app=Flask("My_Excel_App")
@app.route('/home',methods=['GET','POST'])
def myform():
  return render_template('form.html')

@app.route('/output',methods=['GET','POST'])
def output():
  dataset=pd.read_excel("SBI_USER_DB.xlsx",engine='openpyxl')
  #dataset=pd.read_excel("output.xlsx",engine='openpyxl')
  IFSC=request.form.get("IFSC")
  phone=request.form.get("phone")

  if IFSC in list(dataset["IFSC"]):
      if int(phone) not in list(dataset[dataset["IFSC"]==IFSC]["PHONE"]):
        record_index=dataset[dataset["IFSC"]==IFSC].index[0]
        #dataset.get_value(0,'PHONE')= phone
        dataset.at[record_index,'PHONE']= phone
        output_file="C:\\Users\\mohit\\Excel_App\\output.xlsx"
        dataset.to_excel(output_file)
        return "Your Data has been updated to output.xlsx file , check it out... \
                                 Have a Nice Day :)  !!!!  "
      else:
        return "The given Phone NO. already Exists in our Bank Record , \
                   If u want to update ur Phone No. Then plz give Updated Mobile Number ..."
  else:
      return "User does not exist ,  plz provide proper details ..."

  #output_file="C:\\Users\\mohit\\Excel_App\\output.xlsx"
  #dataset.to_excel(output_file)
  #return "Your Data has been updated to output.xlsx file , check it out... \
 #          Have a Nice Day :)  !!!!  "

app.run(port=1234,debug=True)
