# CB-AdES test signature

This repository creates a minimal local CB-AdES baseline signature for the
provided test document. It uses the lowest CB-AdES baseline level, B-B, from
ETSI TS 119 152-1 V1.1.1 (2026-03). This is a generic standards-based test
artifact, not a signature generated for a specific Plugtests portal testcase.

The script creates a tagged COSE_Sign1 object with an attached payload,
ES256 (ECDSA P-256 with SHA-256), the required protected `iat` claim, and a
certificate reference. It generates an ephemeral self-signed test certificate
for each run; the private key is held only in memory and is not written to disk.
The signature is cryptographically verifiable, but the certificate is not
issued by a trusted CA and must not be described as PKI-trusted. No trusted
timestamp or revocation information is included.

## Run

Install the Python dependencies, then generate and locally verify the signature:

```shell
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python tools/create_test_signature.py
```

The output is `generated/CB-AdES-Signature-Test.cbor`. The generated directory
is ignored by Git.

On Windows, run from the repository root:

```powershell
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe tools\create_test_signature.py
```

## Create the kid-friendly code walkthrough

The learning-pack generator creates a short PDF guide, a 30-slide presentation
PDF, and the editable PowerPoint source:

```shell
.venv/bin/python tools/create_learning_pack.py
```

On Windows:

```powershell
.venv\Scripts\python.exe tools\create_learning_pack.py
```

The files are written to `generated/`.
