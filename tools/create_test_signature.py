from __future__ import annotations

from datetime import datetime, timedelta, timezone
import hashlib
from pathlib import Path
import sys

import cbor2
from cryptography import x509
from cryptography.exceptions import InvalidSignature
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import ec
from cryptography.hazmat.primitives.asymmetric.utils import (
    decode_dss_signature,
    encode_dss_signature,
)
from cryptography.x509.oid import NameOID


ROOT = Path(__file__).resolve().parents[1]
PAYLOAD_PATH = ROOT / "test-data" / "CB-AdES-Signature-Test.txt"
SIGNATURE_PATH = ROOT / "generated" / "CB-AdES-Signature-Test.cbor"

COSE_SIGN1_TAG = 18
COSE_ALG_HEADER = 1
COSE_ES256 = -7
COSE_X5T_HEADER = 34
COSE_SHA_256 = -16
COSE_CWT_CLAIMS_HEADER = 15
COSE_IAT_CLAIM = 6
CBADES_UHEADERS = 268
COSE_X5CHAIN = 33


def create_test_certificate(
    private_key: ec.EllipticCurvePrivateKey, now: datetime
) -> x509.Certificate:
    subject = x509.Name(
        [
            x509.NameAttribute(NameOID.COMMON_NAME, "Test Signer"),
            x509.NameAttribute(NameOID.ORGANIZATIONAL_UNIT_NAME, "Student / Observer"),
        ]
    )
    return (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(subject)
        .public_key(private_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before((now - timedelta(minutes=1)).replace(tzinfo=None))
        .not_valid_after((now + timedelta(days=30)).replace(tzinfo=None))
        .add_extension(x509.BasicConstraints(ca=False, path_length=None), critical=True)
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                content_commitment=False,
                key_encipherment=False,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=False,
                crl_sign=False,
                encipher_only=None,
                decipher_only=None,
            ),
            critical=True,
        )
        .sign(private_key, hashes.SHA256())
    )


def create_signature(payload: bytes) -> bytes:
    now = datetime.now(timezone.utc)
    private_key = ec.generate_private_key(ec.SECP256R1())
    certificate = create_test_certificate(private_key, now)
    certificate_der = certificate.public_bytes(serialization.Encoding.DER)

    protected_headers = {
        COSE_ALG_HEADER: COSE_ES256,
        COSE_CWT_CLAIMS_HEADER: {COSE_IAT_CLAIM: int(now.timestamp())},
        COSE_X5T_HEADER: [COSE_SHA_256, hashlib.sha256(certificate_der).digest()],
    }
    protected_bytes = cbor2.dumps(protected_headers, canonical=True)
    unsigned_x5chain = cbor2.dumps(
        {COSE_X5CHAIN: certificate_der}, canonical=True
    )
    signature_structure = cbor2.dumps(
        ["Signature1", protected_bytes, b"", payload], canonical=True
    )

    der_signature = private_key.sign(
        signature_structure, ec.ECDSA(hashes.SHA256())
    )
    r, s = decode_dss_signature(der_signature)
    cose_signature = r.to_bytes(32, "big") + s.to_bytes(32, "big")

    cose_sign1 = cbor2.CBORTag(
        COSE_SIGN1_TAG,
        [
            protected_bytes,
            {CBADES_UHEADERS: [unsigned_x5chain]},
            payload,
            cose_signature,
        ],
    )
    return cbor2.dumps(cose_sign1, canonical=True)


def validate_signature(signature: bytes, expected_payload: bytes) -> None:
    cose_object = cbor2.loads(signature)
    if not isinstance(cose_object, cbor2.CBORTag) or cose_object.tag != COSE_SIGN1_TAG:
        raise ValueError("The signature is not a tagged COSE_Sign1 object.")

    parts = cose_object.value
    if not isinstance(parts, list) or len(parts) != 4:
        raise ValueError("The COSE_Sign1 object has an invalid structure.")

    protected_bytes, unprotected_headers, payload, cose_signature = parts
    if not isinstance(protected_bytes, bytes) or not isinstance(unprotected_headers, dict):
        raise ValueError("The COSE headers have invalid types.")
    if payload != expected_payload:
        raise ValueError("The embedded payload does not match the test document.")
    if not isinstance(cose_signature, bytes) or len(cose_signature) != 64:
        raise ValueError("The ES256 signature must contain 64 bytes.")

    protected_headers = cbor2.loads(protected_bytes)
    if not isinstance(protected_headers, dict):
        raise ValueError("The protected headers are not a CBOR map.")
    if protected_headers.get(COSE_ALG_HEADER) != COSE_ES256:
        raise ValueError("The protected algorithm is not ES256.")

    claims = protected_headers.get(COSE_CWT_CLAIMS_HEADER)
    if not isinstance(claims, dict) or not isinstance(claims.get(COSE_IAT_CLAIM), int):
        raise ValueError("The protected CWT Claims header is missing the iat claim.")

    x5t = protected_headers.get(COSE_X5T_HEADER)
    if (
        not isinstance(x5t, list)
        or len(x5t) != 2
        or x5t[0] != COSE_SHA_256
        or not isinstance(x5t[1], bytes)
    ):
        raise ValueError("The protected x5t certificate reference is invalid.")

    if set(unprotected_headers) != {CBADES_UHEADERS}:
        raise ValueError("The unprotected header must contain only CB-AdES uHeaders.")
    unsigned_headers = unprotected_headers[CBADES_UHEADERS]
    if not isinstance(unsigned_headers, list) or len(unsigned_headers) != 1:
        raise ValueError("The CB-AdES uHeaders array must contain the certificate chain.")
    unsigned_x5chain = cbor2.loads(unsigned_headers[0])
    if not isinstance(unsigned_x5chain, dict):
        raise ValueError("The unsigned x5chain is not a CBOR map.")
    certificate_der = unsigned_x5chain.get(COSE_X5CHAIN)
    if not isinstance(certificate_der, bytes):
        raise ValueError("The signature does not contain its signing certificate.")
    if not hashlib.sha256(certificate_der).digest() == x5t[1]:
        raise ValueError("The protected x5t does not match the embedded certificate.")

    certificate = x509.load_der_x509_certificate(certificate_der)
    public_key = certificate.public_key()
    if (
        not isinstance(public_key, ec.EllipticCurvePublicKey)
        or not isinstance(public_key.curve, ec.SECP256R1)
    ):
        raise ValueError("The signing certificate does not contain a P-256 key.")
    if certificate.subject != certificate.issuer:
        raise ValueError("The test certificate is not self-signed.")

    signature_structure = cbor2.dumps(
        ["Signature1", protected_bytes, b"", payload], canonical=True
    )
    r = int.from_bytes(cose_signature[:32], "big")
    s = int.from_bytes(cose_signature[32:], "big")
    try:
        public_key.verify(
            encode_dss_signature(r, s),
            signature_structure,
            ec.ECDSA(hashes.SHA256()),
        )
        public_key.verify(
            certificate.signature,
            certificate.tbs_certificate_bytes,
            ec.ECDSA(certificate.signature_hash_algorithm),
        )
    except InvalidSignature as exc:
        raise ValueError("The signature or self-signed certificate did not verify.") from exc


def main() -> int:
    try:
        payload = PAYLOAD_PATH.read_bytes()
        signature = create_signature(payload)
        validate_signature(signature, payload)

        SIGNATURE_PATH.parent.mkdir(parents=True, exist_ok=True)
        SIGNATURE_PATH.write_bytes(signature)
        validate_signature(SIGNATURE_PATH.read_bytes(), payload)
    except (OSError, ValueError, cbor2.CBORDecodeError) as exc:
        print(f"Could not create or validate the CB-AdES test signature: {exc}", file=sys.stderr)
        return 1

    print(f"Generated and locally verified: {SIGNATURE_PATH}")
    print("Profile: CB-AdES-B-B; tagged COSE_Sign1; attached payload; ES256.")
    print("Certificate: self-signed test certificate; not PKI-trusted.")
    print("No trusted timestamp or revocation data is included.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
