from models import db, Question, UserTest
import json
import random
from datetime import datetime, timedelta

def get_moscow_time():
    return datetime.utcnow() + timedelta(hours=3)

def init_db(app):
    with app.app_context():
        try:
            db.create_all()
            print("Таблицы созданы")

            if Question.query.count() == 0:
                questions = [
                    {
                        'text': 'Что означает аббревиатура "ИБ"?',
                        'opt1': 'Информационная безопасность',
                        'opt2': 'Интернет безопасность',
                        'opt3': 'Информационная база',
                        'correct': 1
                    },
                    {
                        'text': 'Что такое фишинг?',
                        'opt1': 'Вид компьютерного вируса',
                        'opt2': 'Метод социальной инженерии для получения конфиденциальных данных',
                        'opt3': 'Способ шифрования данных',
                        'correct': 2
                    },
                    {
                        'text': 'Что из перечисленного является самым надежным паролем?',
                        'opt1': '123456',
                        'opt2': 'password',
                        'opt3': 'qW7#pL2@mR9',
                        'correct': 3
                    },
                    {
                        'text': 'Что такое двухфакторная аутентификация?',
                        'opt1': 'Вход по двум разным паролям',
                        'opt2': 'Метод защиты, требующий два способа подтверждения личности',
                        'opt3': 'Два администратора для входа',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое VPN?',
                        'opt1': 'Виртуальная частная сеть для защищенного соединения',
                        'opt2': 'Антивирусная программа',
                        'opt3': 'Тип компьютерного вируса',
                        'correct': 1
                    },
                    {
                        'text': 'Какое действие не рекомендуется делать с конфиденциальными данными?',
                        'opt1': 'Хранить в зашифрованном виде',
                        'opt2': 'Отправлять по незащищенной электронной почте',
                        'opt3': 'Удалять после использования',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое социальная инженерия?',
                        'opt1': 'Создание социальных сетей',
                        'opt2': 'Психологическое манипулирование людьми для получения информации',
                        'opt3': 'Инженерное проектирование',
                        'correct': 2
                    },
                    {
                        'text': 'Как часто рекомендуется менять пароли?',
                        'opt1': 'Раз в 10 лет',
                        'opt2': 'Никогда',
                        'opt3': 'Каждые 30-90 дней',
                        'correct': 3
                    },
                    {
                        'text': 'Что такое вредоносное ПО (malware)?',
                        'opt1': 'Программа для работы с документами',
                        'opt2': 'Программное обеспечение, наносящее вред компьютеру или данным',
                        'opt3': 'Антивирусная программа',
                        'correct': 2
                    },
                    {
                        'text': 'Что делать при подозрении на утечку данных?',
                        'opt1': 'Ничего не сообщать',
                        'opt2': 'Немедленно сообщить руководителю и службе ИБ',
                        'opt3': 'Удалить все данные',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое SSL/TLS сертификат?',
                        'opt1': 'Паспорт пользователя',
                        'opt2': 'Цифровой сертификат для шифрования соединения',
                        'opt3': 'Антивирусная база',
                        'correct': 2
                    },
                    {
                        'text': 'Как понять, что сайт использует защищенное соединение?',
                        'opt1': 'В адресной строке есть замок и "https://"',
                        'opt2': 'Сайт быстро загружается',
                        'opt3': 'На сайте много рекламы',
                        'correct': 1
                    },
                    {
                        'text': 'Что такое DDoS-атака?',
                        'opt1': 'Атака с целью кражи паролей',
                        'opt2': 'Перегрузка сервера большим количеством запросов',
                        'opt3': 'Взлом базы данных',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое "чистый стол"?',
                        'opt1': 'Уборка рабочего места',
                        'opt2': 'Правило, требующее убирать документы с рабочего стола после работы',
                        'opt3': 'Программа для очистки диска',
                        'correct': 2
                    },
                    {
                        'text': 'Какая информация не является конфиденциальной?',
                        'opt1': 'Пароли сотрудников',
                        'opt2': 'Данные банковских карт',
                        'opt3': 'Публичная информация о компании с официального сайта',
                        'correct': 3
                    },
                    {
                        'text': 'Что такое "инцидент информационной безопасности"?',
                        'opt1': 'Плановое обновление ПО',
                        'opt2': 'Событие, нарушающее безопасность информации',
                        'opt3': 'Новый сотрудник в отделе',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое шифрование данных?',
                        'opt1': 'Сжатие данных',
                        'opt2': 'Преобразование данных в нечитаемый вид без ключа',
                        'opt3': 'Копирование данных',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое бэкап (backup)?',
                        'opt1': 'Удаление данных',
                        'opt2': 'Резервное копирование данных',
                        'opt3': 'Антивирусная проверка',
                        'correct': 2
                    },
                    {
                        'text': 'Какой пароль считается наиболее безопасным?',
                        'opt1': 'Короткий пароль из цифр',
                        'opt2': 'Длинный пароль из букв, цифр и спецсимволов',
                        'opt3': 'Словарное слово',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое "безопасный канал связи"?',
                        'opt1': 'Канал с шифрованием данных',
                        'opt2': 'Кабель без помех',
                        'opt3': 'Быстрое интернет-соединение',
                        'correct': 1
                    },
                    {
                        'text': 'Что делать с найденной на рабочем месте флешкой?',
                        'opt1': 'Вставить в компьютер, чтобы проверить содержимое',
                        'opt2': 'Передать в службу ИБ или руководителю',
                        'opt3': 'Выбросить',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое "компрометация учетной записи"?',
                        'opt1': 'Создание новой учетной записи',
                        'opt2': 'Несанкционированное получение доступа к учетной записи',
                        'opt3': 'Блокировка учетной записи',
                        'correct': 2
                    },
                    {
                        'text': 'Какой протокол используется для защищенной передачи данных в интернете?',
                        'opt1': 'HTTP',
                        'opt2': 'HTTPS',
                        'opt3': 'FTP',
                        'correct': 2
                    },
                    {
                        'text': 'Что такое "защита от социальной инженерии"?',
                        'opt1': 'Антивирусная защита',
                        'opt2': 'Обучение сотрудников распознаванию манипуляций',
                        'opt3': 'Фильтрация трафика',
                        'correct': 2
                    },
                    {
                        'text': 'Какое правило является основным для работы с конфиденциальными документами?',
                        'opt1': 'Хранить в открытом доступе',
                        'opt2': 'Не передавать третьим лицам без разрешения',
                        'opt3': 'Делиться с коллегами по интересам',
                        'correct': 2
                    }
                ]

                for q in sample_questions:
                    question = Question(
                        text=q['text'],
                        option1=q['opt1'],
                        option2=q['opt2'],
                        option3=q['opt3'],
                        correct_option=q['correct']
                    )
                    db.session.add(question)

                db.session.commit()
                print(f" Добавлено {len(sample_questions)} тестовых вопросов")
            else:
                print(f" В базе уже есть {Question.query.count()} вопросов")

        except Exception as e:
            print(f" Ошибка при инициализации БД: {e}")
            db.session.rollback()


def get_random_questions(limit=15):
    try:
        # Получаем все вопросы
        all_questions = Question.query.all()

        if len(all_questions) == 0:
            print(" В базе нет вопросов!")
            return []

        if len(all_questions) <= limit:
            return all_questions

        shuffled = random.sample(all_questions, limit)
        return shuffled

    except Exception as e:
        print(f" Ошибка при получении случайных вопросов: {e}")
        return []


def save_test_result(full_name, rank, department, examiner_name, examiner_rank, answers, score, passed):
    try:
        full_name = str(full_name).strip()[:200] if full_name else 'Не указано'
        rank = str(rank).strip()[:100] if rank else ''
        department = str(department).strip()[:100] if department else 'Не указано'
        examiner_name = str(examiner_name).strip()[:200] if examiner_name else 'Не указано'
        examiner_rank = str(examiner_rank).strip()[:100] if examiner_rank else ''

        answers_json = json.dumps(answers, ensure_ascii=False, default=str)

        user_test = UserTest(
            full_name=full_name,
            rank=rank,
            department=department,
            examiner_name=examiner_name,
            examiner_rank=examiner_rank,
            score=int(score) if score else 0,
            total_questions=len(answers),
            passed=bool(passed),
            answers_json=answers_json,
            test_date=datetime.utcnow() + timedelta(hours=3)
        )
        db.session.add(user_test)
        db.session.commit()
        print(f" Результат сохранен (ID: {user_test.id})")
        return user_test.id

    except Exception as e:
        print(f" Ошибка при сохранении результата: {e}")
        db.session.rollback()
        raise

def get_statistics(filters=None):
    try:
        query = UserTest.query

        if filters:
            if filters.get('full_name'):
                query = query.filter(UserTest.full_name.contains(filters['full_name']))
            if filters.get('department'):
                query = query.filter(UserTest.department.contains(filters['department']))

        results = query.order_by(UserTest.test_date.desc()).all()
        print(f" Найдено {len(results)} записей в статистике")
        return results

    except Exception as e:
        print(f" Ошибка при получении статистики: {e}")
        return []


def get_user_test_by_id(test_id):
    try:
        return UserTest.query.get(test_id)
    except Exception as e:
        print(f" Ошибка при получении теста: {e}")
        return None


def delete_test_result(test_id):
    try:
        test = UserTest.query.get(test_id)
        if test:
            db.session.delete(test)
            db.session.commit()
            print(f" Результат ID {test_id} удален")
            return True
        return False
    except Exception as e:
        print(f"️ Ошибка при удалении результата: {e}")
        db.session.rollback()
        return False


def get_questions_count():
    try:
        return Question.query.count()
    except Exception as e:
        print(f" Ошибка при подсчете вопросов: {e}")
        return 0


def add_question(text, option1, option2, option3, correct_option):
    try:
        question = Question(
            text=str(text).strip(),
            option1=str(option1).strip(),
            option2=str(option2).strip(),
            option3=str(option3).strip(),
            correct_option=int(correct_option)
        )
        db.session.add(question)
        db.session.commit()
        print(f" Вопрос добавлен (ID: {question.id})")
        return question.id
    except Exception as e:
        print(f" Ошибка при добавлении вопроса: {e}")
        db.session.rollback()
        return None


def delete_question(question_id):
    try:
        question = Question.query.get(question_id)
        if question:
            db.session.delete(question)
            db.session.commit()
            print(f" Вопрос ID {question_id} удален")
            return True
        return False
    except Exception as e:
        print(f"️ Ошибка при удалении вопроса: {e}")
        db.session.rollback()
        return False