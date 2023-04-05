from mariadb import connect


def get_sql_connection():
    """

    :return: mysqldb (database) or error
    """
    try:
        mysqldb = connect(
            host='sunserver',  # '192.168.0.2'
            port=3306,  # 3306
            user='sunfashion',
            password='Sunfashion@1234',
            database='online_database'
        )
        # print(mysqldb)
    except Exception as e:
        return {'error': e}
    else:
        return {"mysqldb": mysqldb}


print(get_sql_connection())
