import pandas as pd
import matplotlib.pyplot as plt
import io
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, CallbackQueryHandler

# Load the Excel file and get the sheet names (companies)
excel_file = 'Data-test-companies.xlsx'
companies = pd.ExcelFile(excel_file).sheet_names

# Function to generate the buttons (Доход, Расход, Прибыль, КПН, Restart, Quit)
def get_buttons():
    return [
        [InlineKeyboardButton("Доход", callback_data="Доход")],
        [InlineKeyboardButton("Расход", callback_data="Расход")],
        [InlineKeyboardButton("Прибыль", callback_data="Прибыль")],
        [InlineKeyboardButton("КПН", callback_data="КПН")],
        [InlineKeyboardButton("🔄 Restart", callback_data="restart")],
        [InlineKeyboardButton("❌ Quit", callback_data="quit")]
    ]

# Function to start the bot and show the companies menu
async def start(update: Update, context):
    keyboard = [[InlineKeyboardButton(company, callback_data=company)] for company in companies]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    # Send the company selection message or edit if it's a callback
    if update.callback_query:
        await update.callback_query.message.edit_text("Выберите компанию:", reply_markup=reply_markup)
    else:
        await update.message.reply_text("Выберите компанию:", reply_markup=reply_markup)

# Callback function for handling company selection
async def company_selection(update: Update, context):
    query = update.callback_query
    company = query.data

    if company == "restart":
        await restart(update, context)
        return
    elif company == "quit":
        await quit(update, context)
        return

    context.user_data['company'] = company
    
    # Menu for selecting the type of information
    keyboard = get_buttons()
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await query.edit_message_text(f"Компания: {company}\nВыберите тип данных:", reply_markup=reply_markup)

# Function to find the correct column names dynamically
def find_columns_dynamically(df):
    income_col = next((col for col in df.columns if "Доход" in col), None)
    expenses_col = next((col for col in df.columns if "Расход" in col), None)
    profit_col = next((col for col in df.columns if "Прибыль" in col), None)
    tax_col = next((col for col in df.columns if "КПН" in col), None)
    
    if income_col and expenses_col and profit_col and tax_col:
        return {
            'income': income_col,
            'expenses': expenses_col,
            'profit': profit_col,
            'tax': tax_col
        }
    return None

# Function to create a line plot from the data and return the image
def create_line_plot(df, data_type, column_name):
    plt.figure(figsize=(10, 6))
    plt.plot(df['Месяц'], df[column_name], marker='o', linestyle='-', label=data_type)
    plt.title(f"{data_type} по месяцам")
    plt.xlabel("Месяц")
    plt.ylabel(data_type)
    plt.xticks(rotation=45)
    plt.grid(True)
    plt.legend()

    # Save the plot to a BytesIO object
    plot_img = io.BytesIO()
    plt.savefig(plot_img, format='png')
    plt.close()
    plot_img.seek(0)
    
    return plot_img

# Callback function for handling data type selection and sending a table + graph
async def data_type_selection(update: Update, context):
    query = update.callback_query
    data_type = query.data

    if data_type == "restart":
        await restart(update, context)
        return
    elif data_type == "quit":
        await quit(update, context)
        return

    company = context.user_data['company']

    # Load the selected company's data from the sheet
    df = pd.read_excel(excel_file, sheet_name=company)

    # Find the correct column names dynamically
    columns = find_columns_dynamically(df)
    if not columns:
        await query.message.reply_text("Не удалось найти данные для компании.")
        return
    
    # Select the appropriate column based on the data type
    if data_type == "Доход":
        selected_data = df[['Месяц', columns['income']]]
        plot_img = create_line_plot(df, data_type, columns['income'])
    elif data_type == "Расход":
        selected_data = df[['Месяц', columns['expenses']]]
        plot_img = create_line_plot(df, data_type, columns['expenses'])
    elif data_type == "Прибыль":
        selected_data = df[['Месяц', columns['profit']]]
        plot_img = create_line_plot(df, data_type, columns['profit'])
    elif data_type == "КПН":
        selected_data = df[['Месяц', columns['tax']]]
        plot_img = create_line_plot(df, data_type, columns['tax'])

    # Convert the selected data to a string table format
    table_text = selected_data.to_string(index=False)
    
    # Send the table as a message to the user
    await query.message.reply_text(f"Компания: {company}\nТип данных: {data_type}\n\n{table_text}")

    # Send the graph as an image
    await query.message.reply_photo(photo=plot_img)

    # After sending the table and graph, show all the buttons again (Доход, Расход, Прибыль, КПН, Restart, Quit)
    keyboard = get_buttons()
    reply_markup = InlineKeyboardMarkup(keyboard)
    await query.message.reply_text("Что вы хотите сделать дальше?", reply_markup=reply_markup)

# Function to handle the restart command and reset to company selection
async def restart(update: Update, context):
    await start(update, context)

# Function to handle the quit command and reset the bot to initial state (requires user to type /start again)
async def quit(update: Update, context):
    query = update.callback_query
    await query.message.reply_text("Бот завершил работу. Для перезапуска введите /start.")
    context.user_data.clear()

# Main function to run the bot
def main():
    app = Application.builder().token("7544314995:AAGxfntt-aRqwdS7nTYKY2oAMyz5_Fsr1uI").build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(company_selection, pattern='^(' + '|'.join(companies) + '|restart|quit)$'))
    app.add_handler(CallbackQueryHandler(data_type_selection, pattern="^(Доход|Расход|Прибыль|КПН|restart|quit)$"))

    app.run_polling()  # Listen for updates and handle requests

if __name__ == "__main__":
    main()















