import pandas as pd
from os import listdir, path , remove ,close
from sqlalchemy import create_engine

DIR_FILES=path.abspath(path.dirname(__file__))
files=[]

engine=create_engine('sqlite:///All_DB.db',echo=False)
# meta=MetaData(engine)

## to find all .xlsx &  return path of the files
def find_xlsx_filenames( path_to_dir, suffix=".xlsx" ):
    global files
    file_path=[]
    
    filenames = listdir(path_to_dir)
    files= [ filename for filename in filenames if filename.endswith( suffix ) ]
    
    for file_name in files:
        file_path.append(path.abspath(path.join(DIR_FILES,file_name)))
    return file_path

df=[]
Exception_Flag=False

try:
    with pd.ExcelWriter('All_Excels.xlsx') as writer:
        i=0
        for exl_file_at in find_xlsx_filenames(DIR_FILES):
            df.append(pd.read_excel(exl_file_at))
            df[i].to_sql(name=files[i],con=engine,if_exists='append')
            df[i].to_excel(writer,sheet_name=files[i],index=False)
            
            if(i<=len(files)):
                i+=1
                      
except Exception as e:
    Exception_Flag=True
    print(type(e),':',TypeError(e))
   

finally:
    if(Exception_Flag):
        close(0)
        remove('All_Excels.xlsx')
        print("File not Created")
    else:
        print("Successfully created files both in Directory & Database")
 
#3. CONSTRUCT THE DATABASE FOR RESEARCH USE FROM THE DATABASE

# QueryRun=pd.read_sql_query('SELECT * FROM All_Data WHERE Time_Card_Index=59348998',engine)

# print(QueryRun)
