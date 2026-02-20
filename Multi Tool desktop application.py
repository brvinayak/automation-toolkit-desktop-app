# ============================================================

# Automation Toolkit Desktop Application

# Author: Vinayak R B

# Copyright (c) 2026

# This project is created for portfolio demonstration.

# Unauthorized copying, redistribution, or commercial use of this code without permission is strictly prohibited.

# Licensed under CC BY-NC 4.0

# ============================================================


import tkinter as tk
import os
import requests
import pandas as pd
from bs4 import BeautifulSoup

root=tk.Tk()

root.grid_columnconfigure(0, weight=1)

root.geometry("500x400")
root.title("Multi Tool desktop application")

intro_label=tk.Label(root,text="||   Automation Toolkit   || \n 1.Folder Organizer \n 2.Bulk Renamer \n 3.Excel Automation\n 4.Web Scrapper tool \n 5.Weather API Tool",font=("Arial",10))
intro_label.grid(row=0,column=0)

def open_weather_tool():
    status_label.config(text="Status : Opening weather tool...")
    weather_tool=tk.Toplevel(root)
    weather_tool.grab_set()
    weather_tool.geometry("500x400")
    weather_tool.title("Weather API Tool")
    intro_label=tk.Label(weather_tool,text="Welcome To The Weather API Tool\n Just enter your city name and get Weather Report",font=("Comic Sans",10))
    intro_label.grid(row=0, column=1, columnspan=2, pady=10)

    city_label=tk.Label(weather_tool,text="City Name:",font=("Arial",10))
    city_label.grid(row=1,column=0, padx=10, pady=5)

    city_entry=tk.Entry(weather_tool)
    city_entry.grid(row=1,column=1)

    def url_generator():
        return f"https://wttr.in/{city_entry.get()}?format=j1"

    report_label=tk.Label(weather_tool,text="Weather report:",font=("Times New Roman",10))
    report_label.grid(row=2,column=0)

    report=tk.Label(weather_tool,text=" ",font=("Times New Roman",12))
    report.grid(row=4,column=0,columnspan=2,pady=10)


    def get_report():
        if city_entry.get().strip() == "":
            report.config(text="Please enter a city name!!")
            return

        url=url_generator()
        try:
            response=requests.get(url,timeout=5)
            if response.status_code==200:
                data=response.json()
                weather = data["current_condition"][0]
                weather_report = f"""
                City: {city_entry.get()}
                Temperature: {weather['temp_C']} °C
                Humidity: {weather['humidity']} %
                Condition: {weather['weatherDesc'][0]['value']}
                Feels Like: {weather['FeelsLikeC']} °C
                Wind Speed: {weather['windspeedKmph']} Km/Hr
                Observed At: {weather['localObsDateTime']}
                """
                
            elif response.status_code==404:
                weather_report="Page not found!!"

            else:
                weather_report=f"Error:{response.status_code}"

            

        except:
            weather_report="Failed to connect to Weather server.."

        report.config(text=weather_report)

    button=tk.Button(weather_tool,text="Check Weather Report",command=get_report)
    button.grid(row=2,column=1)

def open_folderorg_tool():
    status_label.config(text="Status : Opening Folder Organiser...")

    folderorg_tool=tk.Toplevel(root)
    folderorg_tool.grab_set()
    folderorg_tool.geometry("500x400")

    folderorg_tool.title("Folder Organizer Tool")

    file_path=tk.Label(folderorg_tool,text="Type File Path : ",font=("Arial",10))
    file_path.grid(row=0,column=0)

    file_path_entry=tk.Entry(folderorg_tool)
    file_path_entry.grid(row=0,column=1)

    def fileorg_report():
        filepath=file_path_entry.get()

        result_label=tk.Label(folderorg_tool,text=" ",font=("Arial",10))
        result_label.grid(row=2,column=0,columnspan=2,pady=10)
        
        if filepath.strip() == "":
            result_label.config(text="Please enter the file path!!")
            return
        
        files=os.listdir(filepath)

        images_count=0
        PDF_Doc_count=0
        Word_Doc_count=0
        music_count=0
        video_count=0
        other_format_count=0

        for file in files:
            if not os.path.isfile(os.path.join(filepath, file)):
                continue

            source=os.path.join(filepath,file)

            if  file.endswith((".jpg", ".jpeg", ".png")):
                folder="Image"
                images_count+=1
            elif file.endswith(".pdf"):
                folder="PDF Document"
                PDF_Doc_count+=1
            elif file.endswith((".doc",".docx")):
                folder="Word Document"
                Word_Doc_count+=1
            elif file.endswith(".mp3"):
                folder="Music"
                music_count+=1
            elif file.endswith(".mp4"):
                folder="Video"
                video_count+=1
            else:
                folder="Other Format"
                other_format_count+=1

            dest_folder=os.path.join(filepath,folder)
            
            if (not os.path.exists(dest_folder)):
                os.mkdir(dest_folder)

            destination=os.path.join(dest_folder,file)
            os.rename(source,destination)
        
        
        final_report=f"""
        Images Moved: {images_count}
        PDF Documents Moved: {PDF_Doc_count}
        Word Documents Moved: {Word_Doc_count}
        Musics Moved: {music_count}
        Videos Moved: {video_count}
        Other format files Moved: {other_format_count}
        """

        result_label.config(text=final_report)
            
    confirm_button=tk.Button(folderorg_tool,text="Organize Files",command=fileorg_report,cursor="hand2")
    confirm_button.grid(row=1,column=1,pady=5)

def open_bulkrenamer():
    status_label.config(text="Status : Opening Bulk Renamer Tool...")

    bulkrenamer_tool=tk.Toplevel(root)
    bulkrenamer_tool.grab_set()
    bulkrenamer_tool.geometry("500x400")

    bulkrenamer_tool.title("Bulk Renamer Tool")

    folder_path=tk.Label(bulkrenamer_tool,text="Type Folder Path : ",font=("Arial",10))
    folder_path.grid(row=0,column=0)

    folder_path_entry=tk.Entry(bulkrenamer_tool)
    folder_path_entry.grid(row=0,column=1)

    def bulkrenamer_report():
        folderpath=folder_path_entry.get()

        result_label=tk.Label(bulkrenamer_tool,text=" ",font=("Arial",10))
        result_label.grid(row=2,column=0,columnspan=2,pady=10)

        if folderpath.strip() == "":
            result_label.config(text="Please enter the folder path!!")
            return

        files= os.listdir(folderpath)

        if len(files) == 0:
            result_label.config(text="The folder is empty. No files to rename.")
        else:
            basename_label=tk.Label(bulkrenamer_tool,text="Enter Base Name: ",font=("arial",10))
            basename_label.grid(row=3,column=0,pady=0)

            basename_entry=tk.Entry(bulkrenamer_tool)
            basename_entry.grid(row=3,column=1,pady=0)

            def bulkrename():
                file_count=0
                base_name=basename_entry.get()
                base_count=1
                for file in files:
                    if(os.path.isfile(os.path.join(folderpath, file))==False):
                        continue
                    file_name, file_extension = os.path.splitext(file)
                    new_file_name = f"{base_name}_{base_count}{file_extension}"
                    base_count += 1
                    if os.path.exists(os.path.join(folderpath, new_file_name)):
                        continue
                    os.rename(os.path.join(folderpath,file),os.path.join(folderpath,new_file_name))
                    file_count += 1

                final_report=f"""
                Renamed {file_count} files in the folder.
                Renaming completed
                """
                finalreport_label=tk.Label(bulkrenamer_tool,text=" ",font=("Arial",10))
                finalreport_label.grid(row=5,column=0,columnspan=2,pady=10)


                finalreport_label.config(text=final_report)

                status_label.config(text="Status: Renaming completed")

            rename=tk.Button(bulkrenamer_tool,text="Rename Files",command=bulkrename,cursor="hand2")
            rename.grid(row=4,column=1,pady=5)

            

        
            

    confirm_button=tk.Button(bulkrenamer_tool,text="Fetch Folder",command=bulkrenamer_report,cursor="hand2")
    confirm_button.grid(row=1,column=1,pady=5)

def excel_automation():
    status_label.config(text="Status : Opening Excel Automation Tool...")

    excel_tool=tk.Toplevel(root)
    excel_tool.grab_set()
    excel_tool.geometry("500x400")

    excel_tool.title("Excel Automation Tool")

    filename_label=tk.Label(excel_tool,text="Type File Name : ",font=("Arial",10))
    filename_label.grid(row=0,column=0)

    filename_entry=tk.Entry(excel_tool)
    filename_entry.grid(row=0,column=1)

    def fetch_file():
        filename=filename_entry.get()

        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
            
        if not os.path.exists(filename):
            report_label.config(text="File Not Found !!")
            return
        else:
            report_label.config(text="File Found")

    def automate_excel():
        filename=filename_entry.get()

        if not filename.endswith(".xlsx"):
            filename += ".xlsx"
            
        if not os.path.exists(filename):
            report_label.config(text="File Not Found !!")
            return

        df = pd.read_excel(filename)
        report=f"""
        {df.head()}
        Original Rows : {df.shape[0]}
        Original Columns : {df.shape[1]}
        """
        report_label.config(text=report)
        original_rows=df.shape[0]

        df=df.dropna()
        df=df.drop_duplicates()

        newfile=newfilename_entry.get()

        if not newfile.endswith(".xlsx"):
            newfile += ".xlsx"

        df.to_excel(newfile,index = False)
        columns=df.columns

        finalreport=f"""
        No of Rows after cleaning : {df.shape[0]}
        No of Columns : {df.shape[1]}
        Removed Rows : {original_rows-df.shape[0]}
        Cleaned file saved successfully as: {newfile}
        """
        
        finalreport_label.config(text=finalreport)

        status_label.config(text="Status: Excel automation completed")

    filefetch_button=tk.Button(excel_tool,text="Fetch File",command=fetch_file,cursor="hand2")
    filefetch_button.grid(row=1,column=1,pady=10)


    newfilename_label=tk.Label(excel_tool,text="Type New File Name : ",font=("Arial",10))
    newfilename_label.grid(row=2,column=0)

    newfilename_entry=tk.Entry(excel_tool)
    newfilename_entry.grid(row=2,column=1)

    automate_button=tk.Button(excel_tool,text="Automate Excel File",command=automate_excel,cursor="hand2")
    automate_button.grid(row=3,column=1,pady=10)
    
    report_label=tk.Label(excel_tool,text=" ",font=("Arial",10))
    report_label.grid(row=4,column=0,columnspan=2,pady=10)

    finalreport_label=tk.Label(excel_tool,text=" ",font=("Arial",10))
    finalreport_label.grid(row=5,column=0,columnspan=2,pady=10)
    
def web_scrapping():
    status_label.config(text="Status : Opening Web Scrapping Tool...")

    webscrap_tool=tk.Toplevel(root)
    webscrap_tool.grab_set()
    webscrap_tool.geometry("500x400")

    webscrap_tool.title("Web Scrapping Tool")

    url_label=tk.Label(webscrap_tool,text="Enter Website URL : ",font=("Arial",10))
    url_label.grid(row=0,column=0)

    url_entry=tk.Entry(webscrap_tool)
    url_entry.grid(row=0,column=1)

    def webscrap():
        url=url_entry.get()

        if url.strip()=="":
            scrap_report_label.config(text="Please enter URL")
            return
        
        if not url.startswith(("http://","https://")):
            url="https://"+url
        
        links=set()

        try:
            response=requests.get(url)

            if response.status_code==200:
                print("Page successfully fetched...\n")
                soup=BeautifulSoup(response.text,"html.parser")
                print(f"Website Title:{soup.title.text}\n")
                print(response.text[:200])
                
                for link in soup.find_all("a"):
                    href=link.get("href")
                    if href:
                        links.add(href)
                
                scrap_report=f"""
                Page Successfully Fetched..
                Website Title : {soup.title.text if soup.title else "No title"}

                Total links Found : {len(links)}
                All links are saved to file: links.txt
                """

            elif response.status_code==404:
                scrap_report="Page Not Found.."

            else:
                scrap_report=f"""
                Error:{response.status_code}
                """

        except:
            scrap_report="Error ocurred while fetching the page..."

        with open ("links.txt","w") as f:
            for link in links:
                f.write(link+"\n")
            f.write(f"Total links :{len(links)}")

        scrap_report_label.config(text=scrap_report)

        status_label.config(text="Status: Web Scrappping completed")


    scrap_button=tk.Button(webscrap_tool,text="Collect All The Links",command=webscrap,cursor="hand2")
    scrap_button.grid(row=1,column=1,pady=10)
    
    scrap_report_label=tk.Label(webscrap_tool,text=" ",font=("Arial",10))
    scrap_report_label.grid(row=2,column=0,columnspan=2,pady=10)


    
folder_organizer_button=tk.Button(root,text="Folder Organizer",command=open_folderorg_tool,cursor="hand2")
folder_organizer_button.grid(row=1,column=0,pady=5)
bulk_renamer_button=tk.Button(root,text="Bulk Renamer",command=open_bulkrenamer,cursor="hand2")
bulk_renamer_button.grid(row=2,column=0,pady=5)
excel_automation_button=tk.Button(root,text="Excel Cleaner",command=excel_automation,cursor='hand2')
excel_automation_button.grid(row=3,column=0,pady=5)
web_scrapper_button=tk.Button(root,text="Web Scrapper",command=web_scrapping,cursor="hand2")
web_scrapper_button.grid(row=4,column=0,pady=5)
weather_api_button=tk.Button(root,text="Weather Report",command=open_weather_tool,cursor="hand2")
weather_api_button.grid(row=5,column=0,pady=5)

status_label=tk.Label(root,text="Status : Ready",anchor="w")
status_label.grid(row=6,column=0,sticky="we",pady=20)




root.mainloop()