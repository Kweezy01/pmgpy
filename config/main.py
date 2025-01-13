# main.py
from src.db_utils import get_session, create_tables

def main():
    # Create or update tables according to your schema
    create_tables()

    # Create a session
    session = get_session()

    # Example usage
    users = session.execute("SELECT * FROM users;").fetchall()
    print("Users:", users)

    # Don’t forget to close your session
    session.close()

if __name__ == "__main__":
    main()
