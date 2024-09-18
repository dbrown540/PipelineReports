from src.weeklyreport import TechnoMileBot
import time

def main():
    tm_bot = TechnoMileBot()
    tm_bot.open_tm()
    tm_bot.type_username(username="danny")
    tm_bot.type_password(password="test")
    tm_bot.click_login()
    tm_bot.handle_mfa()
    time.sleep(10)

if __name__ == "__main__":
    main()