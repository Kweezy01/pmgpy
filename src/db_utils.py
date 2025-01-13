"""
db_utils.py

Handles connections to the Aiven MySQL database using SQLAlchemy + PyMySQL.
"""

import os
from sqlalchemy import create_engine
from sqlalchemy import text
from sqlalchemy.orm import sessionmaker
from dotenv import load_dotenv

load_dotenv()  # Load environment variables from .env if present

def get_engine():
    """
    Creates a SQLAlchemy engine for an Aiven MySQL database, potentially with SSL.
    """
    db_host = os.getenv("DB_HOST")
    db_port = os.getenv("DB_PORT", "3306")
    db_user = os.getenv("DB_USER")
    db_pass = os.getenv("DB_PASS")
    db_name = os.getenv("DB_NAME")
    db_ssl_ca = os.getenv("DB_SSL_CA", "")  # Path to CA certificate, if needed

    # Build the base MySQL connection string
    connection_string = f"mysql+pymysql://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}"

    # If you need SSL, pass ssl={'ca': 'path/to/ca.pem'} to 'connect_args'
    if db_ssl_ca:
        ssl_args = {
            "ssl": {
                "ca": db_ssl_ca
            }
        }
        engine = create_engine(connection_string, connect_args=ssl_args, echo=False)
    else:
        # Non-SSL or if Aiven doesn't enforce SSL for your instance
        engine = create_engine(connection_string, echo=False)

    return engine

def get_session():
    """
    Returns a SQLAlchemy session for DB operations.
    """
    engine = get_engine()
    Session = sessionmaker(bind=engine)
    return Session()

def create_tables():
    """
    Example of how you might programmatically create or update tables
    using SQLAlchemy, if you have models defined.
    """
    engine = get_engine()

    # If using SQLAlchemy models, you might import them and do:
    # Base.metadata.create_all(engine)
    #
    # from src.models import Base
    # Base.metadata.create_all(engine)
    #
    # Or you can read a .sql file with raw queries:
    #
    with open("db/schema.sql", "r", encoding="utf-8") as f:
        raw_sql = f.read()
    statements = raw_sql.split(";")  # Split on semicolons if you have multiple statements
    with engine.connect() as conn:
        for statement in raw_sql.split(";"):
            stmt = statement.strip()
            if stmt:
                conn.execute(stmt)

def insert_data(data):
    """
    Insert data into the DB (placeholder).
    In real usage, define models or do raw SQL inserts.
    """
    session = get_session()
    try:
        # Example raw insert:
        # for row in data:
        #     session.execute(
        #         "INSERT INTO my_table (col1, col2) VALUES (:col1, :col2)",
        #         {"col1": row["val1"], "col2": row["val2"]}
        #     )
        # session.commit()
        pass
    finally:
        session.close()
