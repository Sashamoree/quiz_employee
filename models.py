from flask_sqlalchemy import SQLAlchemy
from datetime import datetime

db = SQLAlchemy()


class UserTest(db.Model):
    __tablename__ = 'user_tests'

    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(200), nullable=False)
    rank = db.Column(db.String(100), nullable=True)
    department = db.Column(db.String(100), nullable=False)
    examiner_name = db.Column(db.String(200), nullable=True)
    examiner_rank = db.Column(db.String(100), nullable=True)  #
    test_date = db.Column(db.DateTime, default=datetime.utcnow)
    score = db.Column(db.Integer, default=0)
    total_questions = db.Column(db.Integer, default=15)
    passed = db.Column(db.Boolean, default=False)
    answers_json = db.Column(db.Text, default='[]')

    def __repr__(self):
        return f'<UserTest {self.id}: {self.full_name}>'


class Question(db.Model):
    __tablename__ = 'questions'

    id = db.Column(db.Integer, primary_key=True)
    text = db.Column(db.Text, nullable=False)
    option1 = db.Column(db.Text, nullable=False)
    option2 = db.Column(db.Text, nullable=False)
    option3 = db.Column(db.Text, nullable=False)
    correct_option = db.Column(db.Integer, nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    def __repr__(self):
        return f'<Question {self.id}: {self.text[:50]}>'