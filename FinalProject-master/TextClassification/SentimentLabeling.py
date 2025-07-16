import torch
from datasets import load_dataset
import evaluate

from transformers import BertTokenizer, BertForSequenceClassification, AdamW, get_scheduler
from torch.utils.data import DataLoader
from tqdm.auto import tqdm

# Load dataset
dataset = load_dataset("FinanceInc/auditor_sentiment")

# Check the unique values in the 'label' column to understand its format
print(dataset['train']['label'])

# Map textual labels to numerical values (if the labels are text-based)
label_mapping = {"negative": 0, "neutral": 1, "positive": 2}

def encode_labels(example):
    # Check if the label is a string or a numerical value
    if isinstance(example["label"], str):
        example["label"] = label_mapping[example["label"]]
    return example

# Apply the encoding to the dataset
dataset = dataset.map(encode_labels)

# Split dataset (75% train, 25% test)
train_size = int(0.75 * len(dataset["train"]))
test_size = len(dataset["train"]) - train_size
train_dataset, test_dataset = dataset["train"].select(range(train_size)), dataset["train"].select(range(train_size, len(dataset["train"])))

# Tokenizer setup
tokenizer = BertTokenizer.from_pretrained("bert-base-uncased")

def tokenize_function(example):
    return tokenizer(example["sentence"], padding="max_length", truncation=True)

# Tokenize both train and test datasets
train_dataset = train_dataset.map(tokenize_function, batched=True)
test_dataset = test_dataset.map(tokenize_function, batched=True)

# Format dataset for PyTorch
train_dataset.set_format(type="torch", columns=["input_ids", "attention_mask", "label"])
test_dataset.set_format(type="torch", columns=["input_ids", "attention_mask", "label"])

train_dataloader = DataLoader(train_dataset, batch_size=16, shuffle=True)
test_dataloader = DataLoader(test_dataset, batch_size=16)

# Load model
model = BertForSequenceClassification.from_pretrained("bert-base-uncased", num_labels=3)
device = torch.device("cuda") if torch.cuda.is_available() else torch.device("cpu")
model.to(device)

optimizer = AdamW(model.parameters(), lr=5e-5)
num_epochs = 3
num_training_steps = num_epochs * len(train_dataloader)
lr_scheduler = get_scheduler("linear", optimizer=optimizer, num_warmup_steps=0, num_training_steps=num_training_steps)

# Training loop
# Training loop
progress_bar = tqdm(range(num_training_steps))
model.train()
for epoch in range(num_epochs):
    for batch in train_dataloader:
        batch = {k: v.to(device) for k, v in batch.items()}
        # Pass 'labels' instead of 'label'
        outputs = model(input_ids=batch["input_ids"], attention_mask=batch["attention_mask"], labels=batch["label"])
        loss = outputs.loss
        loss.backward()

        optimizer.step()
        lr_scheduler.step()
        optimizer.zero_grad()
        progress_bar.update(1)


# Evaluation
metric = evaluate.load("accuracy")
model.eval()
for batch in test_dataloader:
    batch = {k: v.to(device) for k, v in batch.items()}
    with torch.no_grad():
        outputs = model(**batch, labels=batch["label"])
    logits = outputs.logits
    predictions = torch.argmax(logits, dim=-1)
    metric.add_batch(predictions=predictions, references=batch["label"])

accuracy = metric.compute()
print(f"Test Accuracy: {accuracy['accuracy']:.4f}")
