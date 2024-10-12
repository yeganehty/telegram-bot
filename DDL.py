import mysql.connector
from config import db_config



def create_database():
    conn = mysql.connector.connect(user=db_config['user'], password=db_config['password'], host=db_config['host'])
    cursor = conn.cursor(dictionary=True)
    SQL_QUERY = "DROP DATABASE IF EXISTS managesalary"
    cursor.execute(SQL_QUERY)
    SQL_QUERY = "CREATE DATABASE IF NOT EXISTS managesalary"
    cursor.execute(SQL_QUERY)
    conn.commit()
    cursor.close()
    conn.close()
    print('Database managesalary created')


def create_personnel_list_table():
    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor(dictionary=True)
    SQL_QUERY = """
        CREATE TABLE IF NOT EXISTS PERSONNEL_LIST (
            `cid`                       BIGINT UNSIGNED NOT NULL PRIMARY KEY,
            `name`                      VARCHAR(100) NOT NULL,
            `last_name`                 VARCHAR(100) NOT NULL,
            `personnel_id`              BIGINT UNSIGNED NOT NULL,
            `personnel_pass`            BIGINT UNSIGNED NOT NULL,
            `job_position`              ENUM('Employee', 'Manager') DEFAULT "Employee",
            `child_count`               SMALLINT UNSIGNED NOT NULL,
            `rate`                      MEDIUMINT UNSIGNED,
            `picture_number`            BIGINT UNSIGNED
             ) ; """
    cursor.execute(SQL_QUERY)
    conn.commit()
    cursor.close()
    conn.close()
    print('PERSONNEL_LIST table created')


def create_timing_table():
    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor(dictionary=True)
    SQL_QUERY = """
        CREATE TABLE IF NOT EXISTS TIMING (
           `cid`          BIGINT UNSIGNED NOT NULL,
           `date`         DATE,
           `start_time`   TIME,
           `end_time`     TIME,
            FOREIGN KEY(`cid`) REFERENCES personnel_list(`cid`)) ;"""
    cursor.execute(SQL_QUERY)
    conn.commit()
    cursor.close()
    conn.close()
    print('TIMING table created')


def create_working_hours_table():
    conn = mysql.connector.connect(**db_config)
    cursor = conn.cursor(dictionary=True)
    SQL_QUERY = """
        CREATE TABLE IF NOT EXISTS working_hours (
            `cid`            BIGINT UNSIGNED NOT NULL,
            `month`          TINYINT UNSIGNED NOT NULL,
            `working_hours`  DECIMAL(7,2) NOT NULL,
            PRIMARY KEY (cid, month),
            FOREIGN KEY (cid) REFERENCES personnel_list(cid)
        ) ;"""
    cursor.execute(SQL_QUERY)
    conn.commit()
    cursor.close()
    conn.close()
    print('working_hours table created')


if __name__ == "__main__":
    create_database()
    create_personnel_list_table()
    create_timing_table()
    create_working_hours_table()



