"""
Test script for Gemma 3 1B conversational resolution check.
Uses the EXACT same prompt and approach as the original Ollama version
that was already proven to work correctly.

Run: python test_resolution_model.py
"""

import time
import torch
from transformers import AutoModelForCausalLM, AutoTokenizer

MODEL_NAME = "google/gemma-3-1b-it"

TEST_CASES = [
    {
        "name": "Clear resolution — password reset",
        "conversation": """[2026-06-18 10:30] support@exlservice.com:
Hi, I have reset your password. Please try logging in with the temporary password 'Temp@1234' and change it after logging in.

[2026-06-18 09:45] user@client.com:
I am unable to login to the portal. It says invalid credentials. Please help.""",
        "sender": "user@client.com",
        "expected": "YES",
    },
    {
        "name": "Not resolved — issue acknowledged but not fixed",
        "conversation": """[2026-06-18 11:00] support@exlservice.com:
Hi, we have received your request and our team is looking into it. We will update you shortly.

[2026-06-18 10:00] user@client.com:
The payment of Rs. 5000 was deducted from my account but the transaction shows as failed. Please refund.""",
        "sender": "user@client.com",
        "expected": "NO",
    },
    {
        "name": "Partially resolved — one question answered, one ignored",
        "conversation": """[2026-06-18 14:00] support@exlservice.com:
Your loan application has been approved and will be disbursed within 3 working days.

[2026-06-18 13:00] user@client.com:
Can you tell me the status of my loan application? Also, what is the interest rate being applied?""",
        "sender": "user@client.com",
        "expected": "NO",
    },
    {
        "name": "Fully resolved — multi-turn conversation",
        "conversation": """[2026-06-18 16:00] support@exlservice.com:
The interest rate on your loan is 12.5% per annum. Is there anything else I can help you with?

[2026-06-18 15:30] user@client.com:
Thank you. What is the interest rate?

[2026-06-18 15:00] support@exlservice.com:
Your loan application has been approved and will be disbursed within 3 working days.

[2026-06-18 14:00] user@client.com:
Can you tell me the status of my loan application? Also, what is the interest rate being applied?""",
        "sender": "user@client.com",
        "expected": "YES",
    },
    {
        "name": "Not resolved — system error still occurring",
        "conversation": """[2026-06-18 12:00] support@exlservice.com:
We are aware of the issue and our technical team is working on it. We apologize for the inconvenience.

[2026-06-18 11:30] user@client.com:
The application keeps crashing whenever I try to generate a report. This has been happening since yesterday.""",
        "sender": "user@client.com",
        "expected": "NO",
    },
    {
        "name": "Resolved — user confirms it is working",
        "conversation": """[2026-06-18 17:00] user@client.com:
Thank you, it is working now!

[2026-06-18 16:30] support@exlservice.com:
I have cleared the cache on your account. Please try again and let me know if the issue persists.

[2026-06-18 16:00] user@client.com:
The dashboard is not loading. It just shows a blank screen.""",
        "sender": "user@client.com",
        "expected": "YES",
    },
    {
        "name": "Not resolved — sender added more info, no reply yet",
        "conversation": """[2026-06-18 10:30] user@client.com:
Also, this is happening on both Chrome and Edge browsers, so it is not a browser issue.

[2026-06-18 10:00] user@client.com:
I cannot access the KYC verification page. It gives a 404 error.""",
        "sender": "user@client.com",
        "expected": "NO",
    },
    {
        "name": "Resolved — API integration fixed",
        "conversation": """[2026-06-18 15:00] support@exlservice.com:
I have updated the API endpoint configuration on our end. The correct endpoint is https://api.exlservice.com/v2/sync. Please update your integration and let us know if this resolves the issue.

[2026-06-18 14:00] user@client.com:
Our integration with your API is failing. We are getting a 404 error on the sync endpoint.""",
        "sender": "user@client.com",
        "expected": "YES",
    },
]


def run_test(model, tokenizer, conversation, sender):
    # Exact same prompt as the original Ollama version
    prompt = f"""Below is a support ticket conversation, shown NEWEST message first. Each message shows the sender's email and their message.

Conversation:
{conversation}

The person who raised this ticket is: {sender}

Question: Has EVERY question or issue raised by {sender} (across all of their messages) been clearly and explicitly answered or fixed by the OTHER person's replies in this conversation?

If even ONE question or issue from {sender} is still unanswered or unresolved, answer NO.
Only answer YES if you can find an explicit answer/fix for EACH thing {sender} asked about.

Answer with ONLY one word: YES or NO."""

    messages = [{"role": "user", "content": prompt}]
    text = tokenizer.apply_chat_template(
        messages,
        tokenize=False,
        add_generation_prompt=True,
    )
    inputs = tokenizer(text, return_tensors="pt")

    t0 = time.time()
    with torch.no_grad():
        outputs = model.generate(
            **inputs,
            max_new_tokens=5,
            do_sample=False,
            pad_token_id=tokenizer.eos_token_id,
        )
    elapsed = time.time() - t0

    response_ids = outputs[0][inputs["input_ids"].shape[1]:]
    response = tokenizer.decode(response_ids, skip_special_tokens=True).strip().upper()
    return response, elapsed


def main():
    print("=" * 60)
    print("Gemma 3 1B — Resolution Check Test (Original Prompt)")
    print("=" * 60)
    print()

    print(f"Loading {MODEL_NAME} from cache...")
    t0 = time.time()
    tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
    model = AutoModelForCausalLM.from_pretrained(
        MODEL_NAME,
        torch_dtype=torch.float16,
        low_cpu_mem_usage=True,
    )
    print(f"Loaded in {time.time()-t0:.1f}s\n")

    passed = failed = 0
    total_time = 0

    for i, tc in enumerate(TEST_CASES, 1):
        response, elapsed = run_test(model, tokenizer, tc["conversation"], tc["sender"])
        total_time += elapsed

        if "YES" in response and "NO" not in response:
            got = "YES"
        elif "NO" in response:
            got = "NO"
        else:
            got = f"UNCLEAR({response!r})"

        correct = got == tc["expected"]
        status = "PASS" if correct else "FAIL"
        passed += correct
        failed += not correct

        print(f"[{i}/{len(TEST_CASES)}] {status} — {tc['name']}")
        print(f"        Expected: {tc['expected']}  |  Got: {got}  |  Time: {elapsed:.1f}s")
        if not correct:
            print(f"        Raw response: {response!r}")
        print()

    print("=" * 60)
    print(f"Results: {passed}/{len(TEST_CASES)} passed  ({failed} failed)")
    print(f"Average inference time: {total_time/len(TEST_CASES):.1f}s per call")
    print("=" * 60)

    if passed >= 7:
        print("\nGemma 3 1B is working correctly — safe to use in track_outlook_mails_com.py")
    else:
        print(f"\n{failed} failures — review above before using in production")


if __name__ == "__main__":
    main()
