import time
import torch
from transformers import AutoModelForCausalLM, AutoTokenizer

MODEL_NAME = "google/gemma-3-1b-it"

print("Loading model (from cache, should be fast)...")
t0 = time.time()
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model = AutoModelForCausalLM.from_pretrained(
    MODEL_NAME,
    torch_dtype=torch.float16,
    low_cpu_mem_usage=True,
)
print(f"Loaded in {time.time()-t0:.1f}s")

prompt = """<start_of_turn>user
Analyze this support ticket conversation (newest message first):

[10:30] support@exlservice.com: I have reset your password. Please try logging in with the temporary password Temp1234.
[09:45] user@client.com: I cannot log in, getting invalid credentials.

Ticket raised by: user@client.com

Did the support team provide a concrete fix? YES or NO only:<end_of_turn>
<start_of_turn>model
"""

inputs = tokenizer(prompt, return_tensors="pt")
t0 = time.time()
with torch.no_grad():
    outputs = model.generate(**inputs, max_new_tokens=5, do_sample=False, pad_token_id=tokenizer.eos_token_id)
elapsed = time.time() - t0

response = tokenizer.decode(outputs[0][inputs["input_ids"].shape[1]:], skip_special_tokens=True)
print(f"Response: {response!r}")
print(f"Time: {elapsed:.1f}s")
