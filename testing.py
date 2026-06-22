"""
Quick speed + accuracy test for Gemma 3 1B with an optimized shorter prompt.
Shorter prompt = fewer input tokens = faster inference.
Run: python test_gemma_speed.py
"""

import time
import torch
from transformers import AutoModelForCausalLM, AutoTokenizer

MODEL_NAME = "google/gemma-3-1b-it"

print("Loading from cache...")
t0 = time.time()
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model = AutoModelForCausalLM.from_pretrained(
    MODEL_NAME,
    torch_dtype=torch.float16,
    low_cpu_mem_usage=True,
)
print(f"Loaded in {time.time()-t0:.1f}s\n")

TEST_CASES = [
    {
        "name": "Clear resolution",
        "conv": "[10:30] support: I have reset your password, please try Temp@1234.\n[09:45] user@client.com: I cannot log in.",
        "sender": "user@client.com",
        "expected": "YES",
    },
    {
        "name": "Only acknowledged",
        "conv": "[11:00] support: We received your request and are looking into it.\n[10:00] user@client.com: Payment deducted but transaction failed, please refund.",
        "sender": "user@client.com",
        "expected": "NO",
    },
    {
        "name": "Still investigating",
        "conv": "[12:00] support: We are aware and our team is working on it.\n[11:30] user@client.com: App crashes when generating a report.",
        "sender": "user@client.com",
        "expected": "NO",
    },
    {
        "name": "User confirms working",
        "conv": "[17:00] user@client.com: Thank you, it works now!\n[16:30] support: I cleared your account cache, please try again.\n[16:00] user@client.com: Dashboard not loading.",
        "sender": "user@client.com",
        "expected": "YES",
    },
    {
        "name": "Partial — one question ignored",
        "conv": "[14:00] support: Your loan is approved, disbursement in 3 days.\n[13:00] user@client.com: What is my loan status? Also what interest rate applies?",
        "sender": "user@client.com",
        "expected": "NO",
    },
    {
        "name": "API fix provided",
        "conv": "[15:00] support: Updated endpoint is https://api.exl.com/v2/sync, please update your integration.\n[14:00] user@client.com: API sync endpoint returns 404.",
        "sender": "user@client.com",
        "expected": "YES",
    },
]

def run(conv, sender):
    # Short, direct prompt — fewer tokens = faster
    prompt = f"""<start_of_turn>user
Support ticket (newest first):
{conv}

Raised by: {sender}

Has the support team fully resolved every issue raised by {sender}? Answer YES or NO only.<end_of_turn>
<start_of_turn>model
"""
    inputs = tokenizer(prompt, return_tensors="pt")
    t0 = time.time()
    with torch.no_grad():
        outputs = model.generate(
            **inputs,
            max_new_tokens=5,
            do_sample=False,
            pad_token_id=tokenizer.eos_token_id,
        )
    elapsed = time.time() - t0
    response = tokenizer.decode(
        outputs[0][inputs["input_ids"].shape[1]:],
        skip_special_tokens=True
    ).strip().upper()
    return response, elapsed, inputs["input_ids"].shape[1]

passed = failed = 0
total_time = 0

print("=" * 55)
for i, tc in enumerate(TEST_CASES, 1):
    raw, elapsed, n_tokens = run(tc["conv"], tc["sender"])
    got = "YES" if "YES" in raw and "NO" not in raw else "NO" if "NO" in raw else f"UNCLEAR({raw})"
    correct = got == tc["expected"]
    status = "PASS" if correct else "FAIL"
    passed += correct
    failed += not correct
    total_time += elapsed
    print(f"[{i}/6] {status} — {tc['name']}")
    print(f"       Expected:{tc['expected']}  Got:{got}  Time:{elapsed:.1f}s  Tokens:{n_tokens}")

print("=" * 55)
print(f"Results : {passed}/6 passed")
print(f"Avg time: {total_time/len(TEST_CASES):.1f}s per call")
print(f"Total   : {total_time:.1f}s")
