import unittest

import cbor2

from tools.create_test_signature import create_signature, validate_signature


class CreateTestSignatureTests(unittest.TestCase):
    def test_signature_verifies_and_rejects_tampering(self) -> None:
        payload = b"CB-AdES test payload"
        signature = create_signature(payload)
        validate_signature(signature, payload)

        with self.assertRaises(ValueError):
            validate_signature(signature, payload + b" altered")

        cose_object = cbor2.loads(signature)
        cose_object.value[3] = cose_object.value[3][:-1] + bytes(
            [cose_object.value[3][-1] ^ 1]
        )
        altered_signature = cbor2.dumps(cose_object, canonical=True)
        with self.assertRaises(ValueError):
            validate_signature(altered_signature, payload)


if __name__ == "__main__":
    unittest.main()
