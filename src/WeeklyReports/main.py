import time
import schedule
from datetime import datetime
from WeeklyReports.weeklyreport import TechnoMileBot, SharePointBot, WeeklyPipelineReportCreator

def weekly_report():
    start_time = time.time()
    
    # Initialize TechnoMileBot and execute the pipeline report generation
    tm_bot = TechnoMileBot()
    tm_filepath = tm_bot.execute()

    # Initialize SharePointBot and authenticate
    sp_bot = SharePointBot()
    conn = sp_bot._auth()

    # Initialize WeeklyPipelineReportCreator to process the data
    mr_bot = WeeklyPipelineReportCreator()
    mr_bot.execute(tm_filepath)  # Use local_path instead of sp_filepath

    # Upload the updated report back to SharePoint
    sp_bot.upload_file(
        conn,
        local_file_path="Weekly Pipeline Report.xlsx",
        sharepoint_folder_url="/sites/RELIBusinessDevelopment/Shared Documents/BD Department Files/Pipeline",
        file_name="Weekly Pipeline Report.xlsx"
    )
    
    end_time = time.time()
    total_time = end_time - start_time
    print(f"Total time taken: {total_time:.2f} seconds")

def job():
    # Get the current time and run the weekly report
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Running weekly report at {current_time}...")
    weekly_report()

if __name__ == "__main__":
    # Schedule the job every Friday at 8:00 AM
    #schedule.every().friday.at("09:15").do(job)

    print("Scheduler started. Waiting for the next scheduled run...")
    """while True:
        schedule.run_pending()
        time.sleep(60)  # Wait for one minute before checking again"""
    job()
