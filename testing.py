"""
Test script for Gemma 3 1B — raw prompt, no chat template (matching Ollama).
Run: python test_resolution_model.py
"""

import time
import torch
from transformers import AutoModelForCausalLM, AutoTokenizer

MODEL_NAME = "google/gemma-3-1b-it"

TEST_CASES = [
    {"name": "Clear resolution — password reset",
     "conversation": "[10:30] support@exlservice.com:\nI have reset your password. Please try Temp@1234.\n\n[09:45] user@client.com:\nI cannot login, getting invalid credentials.",
     "sender": "user@client.com", "expected": "YES"},
    {"name": "Not resolved — only acknowledged",
     "conversation": "[11:00] support@exlservice.com:\nWe received your request and our team is looking into it. We will update you shortly.\n\n[10:00] user@client.com:\nPayment of Rs 5000 deducted but transaction shows failed. Please refund.",
     "sender": "user@client.com", "expected": "NO"},
    {"name": "Partial — one question ignored",
     "conversation": "[14:00] support@exlservice.com:\nYour loan is approved, disbursement in 3 working days.\n\n[13:00] user@client.com:\nWhat is my loan status? Also what interest rate applies?",
     "sender": "user@client.com", "expected": "NO"},
    {"name": "Fully resolved — multi-turn",
     "conversation": "[16:00] support@exlservice.com:\nThe interest rate is 12.5% per annum.\n\n[15:30] user@client.com:\nThank you. What is the interest rate?\n\n[15:00] support@exlservice.com:\nYour loan is approved, disbursement in 3 working days.\n\n[14:00] user@client.com:\nWhat is my loan status? Also what interest rate applies?",
     "sender": "user@client.com", "expected": "YES"},
    {"name": "Not resolved — still investigating",
     "conversation": "[12:00] support@exlservice.com:\nWe are aware and our technical team is working on it.\n\n[11:30] user@client.com:\nApp keeps crashing when generating a report.",
     "sender": "user@client.com", "expected": "NO"},
    {"name": "Resolved — user confirms working",
     "conversation": "[17:00] user@client.com:\nThank you, it is working now!\n\n[16:30] support@exlservice.com:\nI cleared your account cache, please try again.\n\n[16:00] user@client.com:\nDashboard not loading, blank screen.",
     "sender": "user@client.com", "expected": "YES"},
    {"name": "Not resolved — no reply yet",
     "conversation": "[10:30] user@client.com:\nAlso happening on Chrome and Edge, not browser issue.\n\n[10:00] user@client.com:\nCannot access KYC page, getting 404.",
     "sender": "user@client.com", "expected": "NO"},
    {"name": "Resolved — API fix provided",
     "conversation": "[15:00] support@exlservice.com:\nUpdated the endpoint. Correct URL is https://api.exlservice.com/v2/sync. Please update your integration.\n\n[14:00] user@client.com:\nAPI sync endpoint returns 404.",
     "sender": "user@client.com", "expected": "YES"},
]

def run_test(model, tokenizer, conversation, sender):
    # Raw prompt — no chat template wrapping, matching how Ollama sent it
    prompt = f"""Below is a support ticket conversation, shown NEWEST message first. Each message shows the sender's email and their message.

Conversation:
{conversation}

The person who raised this ticket is: {sender}

Question: Has EVERY question or issue raised by {sender} (across all of their messages) been clearly and explicitly answered or fixed by the OTHER person's replies in this conversation?

If even ONE question or issue from {sender} is still unanswered or unresolved, answer NO.
Only answer YES if you can find an explicit answer/fix for EACH thing {sender} asked about.

Answer with ONLY one word: YES or NO."""

    inputs = tokenizer(prompt, return_tensors="pt")
    t0 = time.time()
    with torch.no_grad():
        outputs = model.generate(
            **inputs, max_new_tokens=5, do_sample=False,
            pad_token_id=tokenizer.eos_token_id,
        )
    elapsed = time.time() - t0
    response = tokenizer.decode(
        outputs[0][inputs["input_ids"].shape[1]:], skip_special_tokens=True
    ).strip().upper()
    return response, elapsed

def main():
    print("=" * 60)
    print("Gemma 3 1B — Raw Prompt (matching Ollama)")
    print("=" * 60)
    print()
    print(f"Loading {MODEL_NAME} from cache...")
    t0 = time.time()
    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
    model = AutoModelForCausalLM.from_pretrained(
        MODEL_NAME, torch_dtype=torch.float16, low_cpu_mem_usage=True,
    )
    print(f"Loaded in {time.time()-t0:.1f}s\n")

    passed = failed = 0
    total_time = 0
    for i, tc in enumerate(TEST_CASES, 1):
        raw, elapsed = run_test(model, tokenizer, tc["conversation"], tc["sender"])
        total_time += elapsed
        got = "YES" if "YES" in raw and "NO" not in raw else "NO" if "NO" in raw else f"UNCLEAR({raw!r})"
        correct = got == tc["expected"]
        passed += correct; failed += not correct
        status = "PASS" if correct else "FAIL"
        print(f"[{i}/8] {status} — {tc['name']}")
        print(f"       Expected:{tc['expected']}  Got:{got}  Time:{elapsed:.1f}s")
        if not correct:
            print(f"       Raw: {raw!r}")
        print()

    print("=" * 60)
    print(f"Results: {passed}/8 passed  |  Avg time: {total_time/8:.1f}s")
    print("=" * 60)

if __name__ == "__main__":
    main()
