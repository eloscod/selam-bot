import os
import logging
from dotenv import load_dotenv
from openpyxl import load_workbook
import telebot
from telebot import types
import time

# === Load Environment Variables ===
# Load environment variables from a .env file to securely store sensitive data like BOT_TOKEN
load_dotenv()
BOT_TOKEN = os.getenv('BOT_TOKEN')

# === Logging Setup ===
# Configure logging to track important events and errors for debugging
logging.basicConfig(level=logging.INFO, format='[%(asctime)s] %(levelname)s - %(message)s')

# === Configurations ===
# Define constants for file paths and data range in Excel sheets
BASE_PATH = 'data/'  # Directory where Excel files are stored
DATA_START_ROW = 5   # Starting row for student data in Excel
DATA_END_ROW = 64    # Ending row for student data in Excel

# === Initialize Bot ===
# Initialize the Telegram bot with the token from environment variables
bot = telebot.TeleBot(BOT_TOKEN)

# === Utility Functions ===
# Check if a value can be converted to a float (used for average validation)
def is_number(val):
    try:
        return float(val) if val else False
    except:
        return False

# Safely retrieve a cell value, returning 'N/A' if invalid
def get_value(cell):
    return cell.value if cell and hasattr(cell, 'value') and cell.value is not None else 'N/A'

# Validate the Excel sheet structure based on expected column count
def validate_excel_structure(ws, semester):
    expected_cols = 18 if semester in ['S1', 'S2'] else 17  # S1/S2 have 18 cols, Ave has 17
    if ws.max_column < expected_cols:
        logging.warning(f"Excel file for {ws.title} has fewer columns than expected ({ws.max_column} < {expected_cols})")
        return False
    return True

# === Loading Animation ===
# Display a simple loading animation with dots before settling on "Processing..."
def get_loading_message(chat_id, message_id):
    dots = ["⏳", "⏳.", "⏳..", "⏳..."]
    for i in range(4):
        bot.edit_message_text(chat_id=chat_id, message_id=message_id, text=dots[i])
        time.sleep(0.5)  # Pause for animation effect
    return bot.edit_message_text(chat_id=chat_id, message_id=message_id, text="⏳ Processing...")

# === /start Command Handler ===
# Welcome message with options to check results or view top 3 students
@bot.message_handler(commands=['start'])
def send_welcome(message):
    max_retries = 3  # Number of retry attempts for network issues
    for attempt in range(max_retries):
        try:
            markup = types.InlineKeyboardMarkup()
            markup.add(types.InlineKeyboardButton("✅ Check Results", callback_data='results'))
            markup.add(types.InlineKeyboardButton("✅ View Top 3", callback_data='top3'))
            bot.reply_to(message,
                "🎓 *Welcome to Selam Islamic Elementary School Result Bot* 🎓\n"
                "--------------------------------\n"
                "📚 Official Bot for Selam Islamic Elementary School\n"
                "🌐 Serving Grades 1-6 with Real-Time Results\n"
                "👇 Click below to get started:\n\n"
                "- Check individual student results\n"
                "- View top 3 performing students\n\n"
                "📋 *Format Example:*\n"
                "`1AS11` → Grade 1, Section A, Semester 1, Student 1\n"
                "`1AAve10` → Grade 1, Section A, Average, Student 10",
                parse_mode="Markdown",
                reply_markup=markup
            )
            break
        except telebot.apihelper.ApiException as e:
            logging.error(f"Attempt {attempt + 1} failed: {str(e)}")
            if attempt < max_retries - 1:
                time.sleep(2 ** attempt)  # Exponential backoff: 1, 2, 4 seconds
            else:
                logging.error("Max retries reached. Could not send welcome message.")
                bot.reply_to(message, "🚫 *Error:* Unable to connect to Telegram. Please check your network and try again later.")

# === /help Command Handler ===
# Provide detailed help information with command examples
@bot.message_handler(commands=['help'])
def send_help(message):
    help_text = (
        "🎓 *Selam Islamic Elementary School Result Bot Help* 🎓\n"
        "--------------------------------\n"
        "| Command       | Description                          |\n"
        "|---------------|--------------------------------------|\n"
        "| `/start`      | Display welcome message and options  |\n"
        "| `/help`       | Show this help message               |\n"
        "| *Check Results* | Select grade, section, semester, and enter student number (1-60) |\n"
        "| *View Top 3*  | Select section, semester, and view top 3 students |\n"
        "--------------------------------\n"
        "📋 *Usage Format:*\n"
        "- Example: `1AS11` (Grade 1, Section A, S1, Student 1)\n"
        "- Example: `1AAve10` (Grade 1, Section A, Ave, Student 10)\n"
        "--------------------------------\n"
        "ℹ️ *Notes:*\n"
        "- Ensure correct grade and section files are available.\n"
        "- Use the 'Back' button to navigate.\n"
        "- Contact admin for technical issues."
    )
    bot.reply_to(message, help_text, parse_mode="Markdown")

# === Handle Inline Keyboard Callbacks ===
# Process button clicks for navigation and selections
@bot.callback_query_handler(func=lambda call: True)
def callback_handler(call):
    if call.data == 'results':
        bot.answer_callback_query(call.id)
        prompt_grade_section(call.message, is_top3=False)
    elif call.data == 'top3':
        bot.answer_callback_query(call.id)
        prompt_grade_section(call.message, is_top3=True)
    elif call.data.startswith('grade_'):
        bot.answer_callback_query(call.id)
        grade_section = call.data.replace('grade_', '').replace('top3', '')  # Extract section (e.g., "1A") after removing prefixes
        if call.data.endswith('_back'):
            prompt_grade_section(call.message, is_top3=call.data.startswith('grade_top3'))
        else:
            next_step = prompt_semester if not call.data.startswith('grade_top3') else prompt_top3_semester
            bot.edit_message_text(
                chat_id=call.message.chat.id,
                message_id=call.message.message_id,
                text=f"📋 *Selection Confirmed* 🎓\n"
                     f"✅ Grade {grade_section[0]}, Section {grade_section[1]}\n"
                     f"--------------------------------\n"
                     f"🌟 Please select the semester:",
                reply_markup=next_step(grade_section),
                parse_mode="Markdown"
            )
    elif call.data.startswith(('semester_', 'top3_')) and not call.data.endswith('_back'):
        bot.answer_callback_query(call.id)
        if call.data.startswith('semester_'):
            grade_section, semester = call.data.replace('semester_', '').split('_')
            bot.edit_message_text(
                chat_id=call.message.chat.id,
                message_id=call.message.message_id,
                text=f"📋 *Selection Confirmed* 🎓\n"
                     f"✅ Grade {grade_section[0]}, Section {grade_section[1]}, Semester {semester}\n"
                     f"--------------------------------\n"
                     f"📝 Please enter your student number (1-60):",
                parse_mode="Markdown"
            )
            bot.register_next_step_handler(call.message, lambda msg: process_results(msg, grade_section, semester))
        else:  # top3_
            section, semester = call.data.replace('top3_', '').split('_')
            process_top3(call.message, section, semester)
    elif call.data.endswith('_back'):
        bot.answer_callback_query(call.id)
        if call.data.startswith('semester_'):
            grade_section = call.data.replace('semester__back', '').split('_')[0]
            prompt_grade_section(call.message, is_top3=False)
        elif call.data.startswith('top3_'):
            section = call.data.replace('top3__back', '').split('_')[0]
            prompt_grade_section(call.message, is_top3=True)

# === Helper Functions for Markups ===
# Generate inline keyboard markup for grade and section selection
def get_grade_section_markup(is_top3=False):
    markup = types.InlineKeyboardMarkup(row_width=4)
    sections = ['1A', '1B', '1C', '2A', '2B', '2C', '3A', '3B', '4A', '4B', '5A', '5B', '6A', '6B']
    for i in range(0, len(sections), 4):
        row = sections[i:i + 4]
        buttons = [types.InlineKeyboardButton(f"✅ {sec}", callback_data=f'grade_{"top3" if is_top3 else ""}{sec}') for sec in row]
        if i == 0:  # Add Back button only to the first row
            buttons.append(types.InlineKeyboardButton("⬅️ Back", callback_data=f'grade_{"top3" if is_top3 else ""}{row[0]}_back'))
        markup.row(*buttons[:4])  # Limit to 4 columns, append Back if needed
    return markup

# Generate inline keyboard markup for semester selection
def get_semester_markup(grade_section, is_top3=False):
    markup = types.InlineKeyboardMarkup(row_width=4)
    markup.add(
        types.InlineKeyboardButton("✅ S1", callback_data=f'{"semester_" if not is_top3 else "top3_"}{grade_section}_S1'),
        types.InlineKeyboardButton("✅ S2", callback_data=f'{"semester_" if not is_top3 else "top3_"}{grade_section}_S2'),
        types.InlineKeyboardButton("✅ Ave", callback_data=f'{"semester_" if not is_top3 else "top3_"}{grade_section}_Ave'),
        types.InlineKeyboardButton("⬅️ Back", callback_data=f'{"semester_" if not is_top3 else "top3_"}{grade_section}_back')
    )
    return markup

# === /results and /top3 Step 1: Prompt Grade and Section ===
# Display initial prompt for grade and section selection
def prompt_grade_section(message, is_top3=False):
    bot.reply_to(message,
        f"{'📋 *Select Grade and Section* 🎓' if not is_top3 else '🏆 *Select Section for Top 3* 🎓'}",
        reply_markup=get_grade_section_markup(is_top3),
        parse_mode="Markdown"
    )

# === /results and /top3 Step 2: Prompt Semester ===
# Return semester markup for /results flow
def prompt_semester(grade_section):
    return get_semester_markup(grade_section, is_top3=False)

# Return semester markup for /top3 flow
def prompt_top3_semester(grade_section):
    return get_semester_markup(grade_section, is_top3=True)

# === /results Step 3: Process Student Number and Fetch Results ===
# Handle student number input and display individual results
def process_results(message, grade_section, semester):
    loading_msg = bot.reply_to(message, "⏳ Processing...")  # Send loading message
    try:
        student_no = message.text.strip()

        if not student_no.isdigit() or not (1 <= int(student_no) <= 60):
            bot.reply_to(message, "❌ *Error:* Invalid student number. Please enter a number between 1 and 60.", parse_mode="Markdown")
            bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
            return

        grade = int(grade_section[0])
        section = grade_section[1]

        valid_sections = ['A', 'B', 'C'] if grade in [1, 2] else ['A', 'B']
        if section not in valid_sections:
            bot.reply_to(message, f"❌ *Error:* Invalid section. Use: {', '.join(valid_sections)}.", parse_mode="Markdown")
            bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
            return

        file_path = f"{BASE_PATH}{grade_section}.xlsx"
        wb = load_workbook(file_path, data_only=True)
        ws = wb[semester]

        if not validate_excel_structure(ws, semester):
            bot.reply_to(message, f"⚠️ *Warning:* Invalid Excel structure for {grade_section} - {semester}. Please contact the admin.", parse_mode="Markdown")
            bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
            return

        name_index = 3 if semester in ['S1', 'S2'] else 2
        max_col = 18 if semester in ['S1', 'S2'] else 17

        for row in ws.iter_rows(min_row=DATA_START_ROW, max_row=DATA_END_ROW, max_col=max_col):
            row_no = str(row[1].value).strip() if row[1].value else None
            if row_no == student_no:
                response = (
                    f"📄 *Student Result - {semester}* 🎓\n"
                    f"--------------------------------\n"
                    f"👤 *Student No:* {get_value(row[1])}\n"
                    f"👤 *Name:* {get_value(row[name_index])}\n"
                    f"🔢 *Sex:* {get_value(row[name_index + 1])}\n"
                    f"🎂 *Age:* {get_value(row[name_index + 2])}\n"
                    f"📚 *Subjects:*\n"
                    f" - Amharic: {get_value(row[name_index + 3])}\n"
                    f" - English: {get_value(row[name_index + 4])}\n"
                    f" - Arabic: {get_value(row[name_index + 5])}\n"
                    f" - Maths: {get_value(row[name_index + 6])}\n"
                    f" - E.S: {get_value(row[name_index + 7])}\n"
                    f" - Moral Edu: {get_value(row[name_index + 8])}\n"
                    f" - Art: {get_value(row[name_index + 9])}\n"
                    f" - HPE: {get_value(row[name_index + 10])}\n"
                    f"💡 *Conduct:* {get_value(row[14]) if semester in ['S1', 'S2'] else 'N/A'}\n"
                    f"🧮 *Sum:* {get_value(row[15]) if semester in ['S1', 'S2'] else get_value(row[13])}\n"
                    f"📊 *Average:* {get_value(row[16]) if semester in ['S1', 'S2'] else get_value(row[14])}\n"
                    f"🏅 *Rank:* {get_value(row[17]) if semester in ['S1', 'S2'] else get_value(row[15])}\n"
                    f"📝 *Remark:* {get_value(row[16]) if semester == 'Ave' else 'N/A'}\n"
                    f"--------------------------------\n"
                    f"✅ *Results Displayed Successfully!*"
                )
                bot.reply_to(message, response, parse_mode="Markdown")
                time.sleep(1)  # Delay to ensure result is visible
                bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
                return

        bot.reply_to(message, f"❌ *Error:* Student `{student_no}` not found in {grade_section} - {semester}. Please verify the student number.", parse_mode="Markdown")
        bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)

    except FileNotFoundError:
        bot.reply_to(message, f"📁 *Error:* File for `{grade_section}` not found. Contact admin.", parse_mode="Markdown")
        bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
    except KeyError:
        bot.reply_to(message, f"🗂️ *Error:* Sheet `{semester}` not found in {grade_section}.xlsx.", parse_mode="Markdown")
        bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
    except Exception as e:
        logging.exception("Unexpected error in /results")
        bot.reply_to(message, f"🚫 *Error:* An unexpected issue occurred. Details: {str(e)}. Contact admin for support.")
        bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)

# === /top3 Step 3: Process Top 3 Students ===
# Handle the display of top 3 students based on section and semester
def process_top3(message, section, semester):
    loading_msg = bot.reply_to(message, "⏳ Processing...")  # Send loading message
    try:
        grade = int(section[0])
        sec_code = section[1]
        valid_sections = ['A', 'B', 'C'] if grade in [1, 2] else ['A', 'B']
        if sec_code not in valid_sections:
            bot.reply_to(message, f"❌ *Error:* Invalid section for grade {grade}. Use: {', '.join(valid_sections)}.", parse_mode="Markdown")
            bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
            return
        file_path = f"{BASE_PATH}{section}.xlsx"
        wb = load_workbook(file_path, data_only=True)
        ws = wb[semester]

        if not validate_excel_structure(ws, semester):
            bot.reply_to(message, f"⚠️ *Warning:* Invalid Excel structure for {section} - {semester}. Please contact the admin.", parse_mode="Markdown")
            bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
            return

        students = []
        for row in ws.iter_rows(min_row=DATA_START_ROW, max_row=DATA_END_ROW, max_col=18):
            no, name, avg = row[1].value, row[3].value, row[16].value
            if no and avg and is_number(avg):
                students.append({'no': no, 'name': name, 'average': float(avg)})
        if not students:
            bot.reply_to(message, f"⚠️ *Warning:* No averages found for {section} - {semester}.", parse_mode="Markdown")
            bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
            return
        top3 = sorted(students, key=lambda x: x['average'], reverse=True)[:3]
        response = f"🏆 *Top 3 Students - {section}, {semester}* 🎓\n" + "\n".join([
            f"--------------------------------\n"
            f"{i+1}. 👤 *Name:* {s['name']} (No: {s['no']}, 📊 *Avg:* {s['average']:.1f})"
            for i, s in enumerate(top3)
        ]) + "\n--------------------------------\n✅ *Results Displayed Successfully!*"
        bot.reply_to(message, response, parse_mode="Markdown")
        time.sleep(1)  # Delay to ensure result is visible
        bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
    except FileNotFoundError:
        bot.reply_to(message, f"📁 *Error:* File for `{section}` not found. Contact admin.", parse_mode="Markdown")
        bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
    except KeyError:
        bot.reply_to(message, f"🗂️ *Error:* Sheet `{semester}` not found in {section}.xlsx.", parse_mode="Markdown")
        bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)
    except Exception as e:
        logging.exception("Error in /top3")
        bot.reply_to(message, f"🚫 *Error:* An unexpected issue occurred. Details: {str(e)}. Contact admin for support.")
        bot.delete_message(chat_id=loading_msg.chat.id, message_id=loading_msg.message_id)

# === Run Bot ===
# Start the bot's polling loop to listen for messages
if __name__ == '__main__':
    logging.info("📡 Bot is running...")
    bot.polling()