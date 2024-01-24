import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from pymorphy2 import MorphAnalyzer

nltk.download()


def extract_keywords(text):
    morph = MorphAnalyzer()

    # Токенизация текста
    words = word_tokenize(text)

    # Удаление стоп-слов
    stop_words = set(stopwords.words("russian"))
    filtered_words = [word for word in words if word.lower() not in stop_words]

    # Лемматизация слов
    lemmas = [morph.parse(word)[0].normal_form for word in filtered_words]

    # Создание текста с тегами для ключевых слов
    tagged_text = text
    for lemma in lemmas:
        tagged_text = tagged_text.replace(lemma, f"<keyword>{lemma}</keyword>")

    return tagged_text


# Пример текста на русском
text = "Обработка естественного языка - это область искусственного интеллекта, которая занимается анализом, интерпретацией и генерацией человеческого языка."
tagged_text = extract_keywords(text)
print(tagged_text)
