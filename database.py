# https://www.youtube.com/watch?v=pd-0G0MigUA
# https://www.tutorialspoint.com/sqlite/sqlite_python.htm
# https://blog.miguelgrinberg.com/post/restful-authentication-with-flask

from passlib.apps import custom_app_context as pwd_context
import sqlite3

database_path = 'data/database.db'


def initialise():
    conn = sqlite3.connect(database_path)
    conn.execute("""   CREATE TABLE CREDENTIALS
                        (USERNAME       TEXT PRIMARY KEY             NOT NULL,
                        PASSWORD        TEXT                         NOT NULL);""")


def insert(username, password):
    conn = sqlite3.connect(database_path)
    conn.execute("INSERT INTO CREDENTIALS (USERNAME, PASSWORD) \
                VALUES ('" + username + "', '" + pwd_context.hash(password) + "')")
    conn.commit()
    conn.close()


def exist(username, password):
    conn = sqlite3.connect(database_path)
    cursor = conn.execute("  SELECT PASSWORD \
                             FROM  CREDENTIALS \
                             WHERE USERNAME = '" + username + "'")
    row = cursor.fetchone()
    conn.close()

    if row is None:
        return False

    return pwd_context.verify(password, row[0])


def print_all_rows():
    conn = sqlite3.connect(database_path)
    cursor = conn.execute(" SELECT * \
                            FROM CREDENTIALS")
    print(cursor.fetchall())
    conn.close()


def delete(username):
    conn = sqlite3.connect(database_path)
    conn.execute("  DELETE \
                    FROM CREDENTIALS \
                    WHERE USERNAME = '" + username + "'")
    conn.commit()
    conn.close()


if __name__ == '__main__':
    # initialise()
    # insert("username", "password")
    # delete("username")
    print_all_rows()
    # print(exist('username', 'password'))
