from transformers import AutoTokenizer

tokenizer = AutoTokenizer.from_pretrained(
    "sentence-transformers/all-MiniLM-L6-v2"
)

text = "Data Engnieering"

tokens = tokenizer.tokenize(text)

print(tokens)