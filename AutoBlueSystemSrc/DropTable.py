import sqlite3

def drop_orders_table(db_path):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    cursor.execute("DROP TABLE IF EXISTS orders")
    
    conn.commit()
    conn.close()

if __name__ == "__main__":
    db_path = 'orders.db'  # Specify your database path
    drop_orders_table(db_path)
    print("Orders table has been dropped.")
