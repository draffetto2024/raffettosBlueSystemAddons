import sqlite3
import pyautogui
import time
from collections import defaultdict
import keyboard  # For handling keyboard interrupts

# Flag to indicate if the process should stop
stop_typing = False

def get_orders():
    conn = sqlite3.connect('orders.db')  # Replace with your database file
    cursor = conn.cursor()
    cursor.execute("SELECT customer, customer_id, cases, lbs, item, item_id, date_generated, date_processed FROM orders")
    rows = cursor.fetchall()
    conn.close()

    orders = defaultdict(list)
    for row in rows:
        customer = row[0]
        customer_id = row[1]
        orders[(customer, customer_id)].append(row[2:])  # Group by customer and customer_id
    return orders

def enter_order(customer, customer_id, order_details):
    global stop_typing
    # Simulate keystrokes for the order entry
    pyautogui.typewrite(f'Customer: {customer}\n', interval=0.1)
    pyautogui.typewrite(f'Customer ID: {customer_id}\n', interval=0.1)
    for detail in order_details:
        if stop_typing:  # Check if the 'esc' key is pressed to interrupt
            print("Process interrupted by user.")
            return False
        cases, lbs, item, item_id, date_generated, date_processed = detail
        pyautogui.typewrite(f'Cases: {cases}\n', interval=0.1)
        pyautogui.typewrite(f'Lbs: {lbs}\n', interval=0.1)
        pyautogui.typewrite(f'Item: {item}\n', interval=0.1)
        pyautogui.typewrite(f'Item ID: {item_id}\n', interval=0.1)
        pyautogui.typewrite(f'Date Generated: {date_generated}\n', interval=0.1)
        pyautogui.typewrite(f'Date Processed: {date_processed}\n', interval=0.1)
        pyautogui.press('enter')
    pyautogui.press('enter')  # Additional enter to separate orders if needed
    return True

def on_esc_pressed(e):
    global stop_typing
    stop_typing = True

if __name__ == "__main__":
    # Set up the interrupt listener
    keyboard.on_press_key("esc", on_esc_pressed)

    # Get orders from the database
    orders = get_orders()

    # Wait for the user to switch to the correct tab
    print("Please switch to the desired tab within 5 seconds.")
    print("Hold 'esc' to interrupt the process.")
    time.sleep(5)  # Give the user 5 seconds to switch to the correct tab

    # Enter each order
    for (customer, customer_id), order_details in orders.items():
        if not enter_order(customer, customer_id, order_details):
            break
        time.sleep(1)  # Add a delay between orders if necessary
