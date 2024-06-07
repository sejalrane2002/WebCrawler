import requests
from bs4 import BeautifulSoup
import nltk
import re
import pandas as pd
from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.corpus import stopwords

# Load NLTK resources
nltk.download('punkt')
nltk.download('stopwords')

# Function to extract article text from URL
def extract_text_from_url(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        html_content = response.text
        soup = BeautifulSoup(html_content, 'html.parser')
        article_title = soup.find('title').get_text()
        paragraphs = soup.find_all('p')
        article_text = "\n".join([paragraph.get_text() for paragraph in paragraphs])
        return article_title, article_text
    except Exception as e:
        print(f"Error extracting text from {url}: {e}")
        return None, None

# Function to perform sentiment analysis
def sentiment_analysis(text):
    try:
        with open(r'C:\\Users\\Sejal\\Desktop\\task\\MasterDictionary\\positive-words.txt', 'r') as f:
            positive_words = set(f.read().splitlines())
        with open(r'C:\\Users\\Sejal\\Desktop\\task\\MasterDictionary\\negative-words.txt', 'r') as f:
            negative_words = set(f.read().splitlines())

        tokens = word_tokenize(text)
        stop_words = set(stopwords.words('english'))
        filtered_tokens = [word for word in tokens if word.lower() not in stop_words]

        positive_score = sum(1 for word in filtered_tokens if word.lower() in positive_words)
        negative_score = sum(1 for word in filtered_tokens if word.lower() in negative_words)

        polarity_score = (positive_score - negative_score) / ((positive_score + negative_score) + 0.000001)
        subjectivity_score = (positive_score + negative_score) / (len(filtered_tokens) + 0.000001)

        return positive_score, negative_score, polarity_score, subjectivity_score
    except Exception as e:
        print(f"Error performing sentiment analysis: {e}")
        return 0, 0, 0, 0

# Function to perform readability analysis
def readability_analysis(text):
    try:
        sentences = sent_tokenize(text)
        word_tokens = word_tokenize(text)
        word_count = len(word_tokens)
        avg_sentence_length = word_count / len(sentences)

        complex_word_count = sum(1 for word in word_tokens if len(re.findall(r'[aeiouy]', word.lower())) > 2)
        percentage_of_complex_words = (complex_word_count / word_count) * 100

        fog_index = 0.4 * (avg_sentence_length + percentage_of_complex_words)

        syllable_per_word = sum(len(re.findall(r'[aeiouy]', word.lower())) for word in word_tokens) / word_count
        personal_pronouns_count = len(re.findall(r'\b(I|we|my|ours|us)\b', text, re.IGNORECASE))
        avg_word_length = sum(len(word) for word in word_tokens) / word_count

        return avg_sentence_length, percentage_of_complex_words, fog_index, avg_sentence_length, \
               complex_word_count, word_count, syllable_per_word, personal_pronouns_count, avg_word_length
    except Exception as e:
        print(f"Error performing readability analysis: {e}")
        return 0, 0, 0, 0, 0, 0, 0, 0, 0

# Main function to iterate through URLs, extract data, and generate output
def main():
    try:
        input_df = pd.read_excel(r'C:\\Users\\Sejal\\Desktop\\task\\Input.xlsx')
        output_data = []

        for index, row in input_df.iterrows():
            url = row['URL']
            url_id = row['URL_ID']

            article_title, article_text = extract_text_from_url(url)
            if article_text is None:
                continue

            positive_score, negative_score, polarity_score, subjectivity_score = sentiment_analysis(article_text)
            avg_sentence_length, percentage_of_complex_words, fog_index, avg_words_per_sentence, \
            complex_word_count, word_count, syllable_per_word, personal_pronouns_count, avg_word_length = readability_analysis(article_text)

            output_data.append([url_id, article_title, positive_score, negative_score, polarity_score, subjectivity_score,
                                avg_sentence_length, percentage_of_complex_words, fog_index, avg_words_per_sentence,
                                complex_word_count, word_count, syllable_per_word, personal_pronouns_count, avg_word_length])

            print(f"Processed URL_ID: {url_id}")

        columns = ['URL_ID', 'Article_Title', 'POSITIVE_SCORE', 'NEGATIVE_SCORE', 'POLARITY_SCORE', 'SUBJECTIVITY_SCORE',
                   'AVG_SENTENCE_LENGTH', 'PERCENTAGE_OF_COMPLEX_WORDS', 'FOG_INDEX', 'AVG_NUMBER_OF_WORDS_PER_SENTENCE',
                   'COMPLEX_WORD_COUNT', 'WORD_COUNT', 'SYLLABLE_PER_WORD', 'PERSONAL_PRONOUNS', 'AVG_WORD_LENGTH']
        output_df = pd.DataFrame(output_data, columns=columns)

        output_df.to_excel(r'C:\\Users\\Sejal\\Desktop\\task\\Output.xlsx', index=False)
        print("Data successfully written to Output.xlsx")

    except Exception as e:
        print(f"Error in main function: {e}")

if __name__ == "__main__":
    main()
