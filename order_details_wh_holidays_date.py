from datetime import datetime, timedelta
import psycopg2

conn = psycopg2.connect(
     host='borosilcp.postgres.database.azure.com',
     database='easyecom',
     user='db_user',
     password='Hello123!'
)
cursor = conn.cursor()
cursor.execute("SELECT order_date, assigned_warehouse_id FROM order_details WHERE order_date between '2024-03-23 00:00:00' and '2024-03-30 23:59:59'; ")
order_dates = cursor.fetchmany(size=50)

cursor.execute("SELECT date, wh_id FROM wh_holidays")
holidays = cursor.fetchall()


def get_next_valid_date(order_date, wh_id, holidays):
    order_date = datetime.strptime(order_date, '%Y-%m-%d %H:%M:%S').date()
    
    next_valid_date = order_date + timedelta(days=1)        
    while next_valid_date.weekday() == 5 or next_valid_date.weekday() == 6 or (next_valid_date, wh_id) in holidays:
            next_valid_date += timedelta(days=1)
        
                         
    return order_date.strftime('%A'), next_valid_date.strftime('%Y-%m-%d %H:%M:%S') 

for order_date, wh_id in order_dates:
    days_name, next_valid_date= get_next_valid_date(order_date, wh_id, holidays)
    print(f"Order Date: {order_date} ({days_name}), Next Valid Date: {next_valid_date}, WH ID: {wh_id}")

cursor.close()
conn.close()