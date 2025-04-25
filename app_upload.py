import fileinput
from flask import Flask, render_template, request,send_from_directory
import os
import Docx_Converter
import Excel_Converter
from zipfile import ZipFile

app = Flask(__name__)

dir ="C:\CBSE"

if not os.path.exists(dir):
    os.mkdir(dir)

app.config["UPLOAD_PATH"] = dir

@app.route('/',methods=['GET','POST'])
def upload_file():
    if request.method == 'POST':
        f = request.files['file_name']
        lst = request.files['list_name']
        f_name = dir + "\\" + f.filename
        lst_name = dir + "\\" + lst.filename
        f.save(f_name)
        if lst_name != dir + "\\":
            lst.save(lst_name)
        select = request.form.get("Class")
        select1 = request.form.get("Function")
        select2 = request.form.get("mode")
        if select == 'Grade 10' and select1 == "Marksheet" and select2 == "Both":
            Excel_Converter.convertG10(f_name)
            return send_from_directory("C:\CBSE",f.filename[:-4]+".csv")
        elif select == 'Grade 10' and select1 == "Marksheet" and select2 == "Dayscholar/Border":
            Excel_Converter.convertG10(f_name,lst_name)
            obj=ZipFile('C:\CBSE\\results.zip','w')
            obj.write("C:\CBSE\\" + f.filename[:-4] + "(DS).csv")
            obj.write("C:\CBSE\\" + f.filename[:-4] + "(BD).csv")
            obj.close()
            os.remove("C:\CBSE\\" + f.filename[:-4] + "(DS).csv")
            os.remove("C:\CBSE\\" + f.filename[:-4] + "(BD).csv")
            return send_from_directory('C:\CBSE','results.zip')
        elif select == 'Grade 10' and select1 == "Subject Toppers" and select2 == "Both":
            Docx_Converter.sub_top_g10(f_name)
            return send_from_directory("C:\CBSE",f.filename[:-4]+".docx")
        elif select == 'Grade 10' and select1 == "Subject Toppers" and select2 == "Dayscholar/Border":
            Docx_Converter.sub_top_g10(f_name,lst_name)
            obj = ZipFile('C:\CBSE\\results.zip', 'w')
            obj.write("C:\CBSE\\" + f.filename[:-4] + "(DS).docx")
            obj.write("C:\CBSE\\" + f.filename[:-4] + "(BD).docx")
            obj.close()
            os.remove("C:\CBSE\\" + f.filename[:-4] + "(DS).docx")
            os.remove("C:\CBSE\\" + f.filename[:-4] + "(BD).docx")
            return send_from_directory('C:\CBSE', 'results.zip')
        elif select == 'Grade 12' and select1 == "Marksheet" and select2 == "Both":
            Excel_Converter.convertG12(f_name)
            return send_from_directory("C:\CBSE",f.filename[:-4]+".csv")
        elif select == 'Grade 12' and select1 == "Marksheet" and select2 == "Dayscholar/Border":
            Excel_Converter.convertG12(f_name,lst_name)
            obj=ZipFile('C:\CBSE\\results.zip','w')
            obj.write("C:\CBSE\\" + f.filename[:-4] + "(DS).csv")
            obj.write("C:\CBSE\\" + f.filename[:-4] + "(BD).csv")
            obj.close()
            os.remove("C:\CBSE\\" + f.filename[:-4] + "(DS).csv")
            os.remove("C:\CBSE\\" + f.filename[:-4] + "(BD).csv")
            return send_from_directory('C:\CBSE','results.zip')
        elif select == 'Grade 12' and select1 == "Subject Toppers" and select2 == "Both":
            Docx_Converter.sub_top_g12(f_name)
            return send_from_directory("C:\CBSE",f.filename[:-4]+".docx")
        elif select == 'Grade 12' and select1 == "Subject Toppers" and select2 == "Dayscholar/Border":
            Docx_Converter.sub_top_g12(f_name,lst_name)
            obj = ZipFile('C:\CBSE\\results.zip', 'w')
            obj.write("C:\CBSE\\" + f.filename[:-4] + "(DS).docx")
            obj.write("C:\CBSE\\" + f.filename[:-4] + "(BD).docx")
            obj.close()
            os.remove("C:\CBSE\\" + f.filename[:-4] + "(DS).docx")
            os.remove("C:\CBSE\\" + f.filename[:-4] + "(BD).docx")
            return send_from_directory('C:\CBSE', 'results.zip')
        else:
            print("Error")

        os.remove(f_name)
        if lst_name != dir + "\\":
            os.remove(lst_name)
            os.remove('C:\CBSE\\results.zip')
        else:
            os.remove("C:\CBSE" + f.filename[:-4] + ".csv")

    return render_template('upload.html')

if __name__ == '__main__':
    app.run(debug=True)