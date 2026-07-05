import mysql.connector
from db_config import *

def get_connection():

    connection = mysql.connector.connect(

        host=DB_HOST,
        user=DB_USER,
        password=DB_PASSWORD,
        database=DB_NAME

    )

    return connection