from __future__ import annotations

from pathlib import Path
import sys
from xml.sax.saxutils import escape

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.pagesizes import landscape, letter
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "generated"
GUIDE_PDF = OUTPUT / "CB-AdES-Code-Adventure-Guide.pdf"
SLIDES_PDF = OUTPUT / "CB-AdES-Code-Adventure-30-Slides.pdf"
POWERPOINT = OUTPUT / "CB-AdES-Code-Adventure-30-Slides.pptx"

NAVY = "#183153"
INK = "#26364A"
MUTED = "#53657A"
PAPER = "#F5F8FC"
WHITE = "#FFFFFF"
TEAL = "#26A69A"
ORANGE = "#F29F05"
PINK = "#E76F8A"
BLUE = "#4D7CFE"
PURPLE = "#805AD5"
GREEN = "#3D9970"
PALETTE = [TEAL, ORANGE, PINK, BLUE, PURPLE, GREEN]

SLIDES = [
    {
        "title": "The Tiny Signature Machine",
        "subtitle": "A friendly tour of the code that packs a message into a digital signature.",
        "bullets": [
            "We will follow one little text file on a big adventure.",
            "The program makes a temporary key, signs the message, then checks its work.",
            "Think of the signature as a tamper-evident seal—not a magic trust badge.",
        ],
        "callout": "A 30-slide code adventure for curious young builders!",
        "badge": "WELCOME!",
    },
    {
        "title": "What does this project do?",
        "subtitle": "It makes a signed envelope around a small text message.",
        "bullets": [
            "The message is the payload: the words we want to protect.",
            "A digital signature helps detect whether those words were changed.",
            "A certificate carries a public key and a name label for this test.",
        ],
        "callout": "It is a local learning demo, not an official ETSI signature.",
        "badge": "THE MISSION",
    },
    {
        "title": "Meet the project folders",
        "subtitle": "Each folder has one job, like rooms in a workshop.",
        "bullets": [
            "test-data/ holds the sample text document.",
            "tools/ contains the signature maker and this presentation maker.",
            "tests/ checks that good signatures pass and changed ones fail.",
        ],
        "callout": "generated/ is where the program puts its results.",
        "badge": "WORKSHOP MAP",
    },
    {
        "title": "How do we start the machine?",
        "subtitle": "Python runs the instructions from top to bottom.",
        "bullets": [
            "A virtual environment keeps this project's packages together.",
            "requirements.txt lists the helpers Python needs.",
            "Then we run tools/create_test_signature.py.",
        ],
        "callout": "The README has copy-and-paste commands for Mac and Windows.",
        "badge": "START HERE",
    },
    {
        "title": "The helper toolbox",
        "subtitle": "The imports at the top bring useful tools into the program.",
        "bullets": [
            "cbor2 packs and unpacks compact CBOR data.",
            "cryptography creates keys and certificates, then signs and verifies.",
            "hashlib makes a short fingerprint of the certificate.",
        ],
        "callout": "The code reuses trusted libraries instead of inventing crypto.",
        "badge": "IMPORTS",
    },
    {
        "title": "Labels on the packages",
        "subtitle": "Constants give important COSE and CB-AdES numbers readable names.",
        "bullets": [
            "COSE_SIGN1_TAG is 18: the envelope's CBOR tag.",
            "COSE_ES256 is -7: ECDSA on P-256 with SHA-256.",
            "Other labels identify the certificate, time, and headers.",
        ],
        "callout": "A named label is easier to understand than a mystery number.",
        "badge": "CONSTANTS",
    },
    {
        "title": "Our sample message",
        "subtitle": "The program reads the sample text as bytes.",
        "bullets": [
            "It lives in test-data/CB-AdES-Signature-Test.txt.",
            "The message says this is a CB-AdES test by a student observer.",
            "The signature protects these exact bytes, including line breaks.",
        ],
        "callout": "Change even one character and the old signature should fail.",
        "badge": "THE PAYLOAD",
    },
    {
        "title": "What time is it?",
        "subtitle": "The signer records a claimed signing time in UTC.",
        "bullets": [
            "UTC is a shared clock standard used around the world.",
            "The program stores the time as a number of seconds.",
            "This is a claimed time—not a trusted timestamp from a time server.",
        ],
        "callout": "A clock note is not the same as a TSA time-stamp.",
        "badge": "TIME NOTE",
    },
    {
        "title": "Make a one-time key",
        "subtitle": "A fresh P-256 elliptic-curve key pair is created in memory.",
        "bullets": [
            "The private key is the secret signing tool.",
            "The matching public key can check the signature.",
            "This demo does not save the private key to a file.",
        ],
        "callout": "Keep real private keys secret; never paste them into chat.",
        "badge": "KEY MAKER",
    },
    {
        "title": "Build a test certificate",
        "subtitle": "A certificate is like an ID card that carries a public key.",
        "bullets": [
            "Its subject says Test Signer, Student / Observer.",
            "It includes the public half of the new key pair.",
            "It is created for testing and signs itself.",
        ],
        "callout": "A self-signed test ID is not verified by a trusted authority.",
        "badge": "TEST ID",
    },
    {
        "title": "Set certificate dates",
        "subtitle": "The certificate gets a serial number and a short lifetime.",
        "bullets": [
            "It starts about one minute before the current time.",
            "It expires about 30 days later.",
            "The random serial number helps distinguish this certificate.",
        ],
        "callout": "Short-lived test certificates are not long-term identity cards.",
        "badge": "CERTIFICATE",
    },
    {
        "title": "Give the key one job",
        "subtitle": "Certificate extensions describe what the test key may do.",
        "bullets": [
            "BasicConstraints says it is not a certificate authority.",
            "KeyUsage allows making digital signatures.",
            "It is not marked for issuing other certificates.",
        ],
        "callout": "The code signs the certificate with SHA-256 too.",
        "badge": "KEY RULES",
    },
    {
        "title": "A certificate signs itself",
        "subtitle": "The certificate is signed using its own private key.",
        "bullets": [
            "That makes the certificate self-signed.",
            "Later, the validator checks this certificate signature.",
            "That check proves the math matches—not that an authority trusts it.",
        ],
        "callout": "Math can verify a signature while trust is still missing.",
        "badge": "SELF-SIGNED",
    },
    {
        "title": "Turn the certificate into bytes",
        "subtitle": "DER is a compact binary way to write a certificate.",
        "bullets": [
            "The program converts the certificate to DER bytes.",
            "Those bytes can be placed inside the CBOR envelope.",
            "The certificate contains public information, not the private key.",
        ],
        "callout": "Binary data is compact but not meant to read like a paragraph.",
        "badge": "DER BYTES",
    },
    {
        "title": "Three clues go in the envelope",
        "subtitle": "Protected headers travel with the signed message.",
        "bullets": [
            "alg tells the verifier which signature algorithm was used.",
            "CWT Claims carries the signing-time claim, iat.",
            "x5t carries a hash fingerprint of the signing certificate.",
        ],
        "callout": "These clues are protected by the digital signature.",
        "badge": "HEADERS",
    },
    {
        "title": "Fingerprint the certificate",
        "subtitle": "A hash turns certificate bytes into a short, fixed-size fingerprint.",
        "bullets": [
            "The code computes SHA-256 over the DER certificate.",
            "The fingerprint helps check which certificate is referenced.",
            "The validator compares it with the certificate inside the envelope.",
        ],
        "callout": "A hash is like a tamper-sensitive checksum, not encryption.",
        "badge": "FINGERPRINT",
    },
    {
        "title": "Pack protected headers",
        "subtitle": "The header map is encoded as canonical CBOR bytes.",
        "bullets": [
            "Canonical encoding gives the same data a consistent byte layout.",
            "The encoded header bytes become part of what is signed.",
            "Changing a protected header changes the signature check.",
        ],
        "callout": "CBOR is a compact format for structured data.",
        "badge": "CBOR PACKER",
    },
    {
        "title": "Prepare the signing recipe",
        "subtitle": "COSE defines the exact pieces that go into the signature.",
        "bullets": [
            "The context says Signature1.",
            "The recipe includes protected bytes and the attached payload.",
            "An empty byte string here means there is no external AAD.",
        ],
        "callout": "Signer and verifier must rebuild exactly the same recipe.",
        "badge": "SIGNATURE1",
    },
    {
        "title": "Make the ECDSA signature",
        "subtitle": "The private key signs the recipe using ECDSA and SHA-256.",
        "bullets": [
            "First, SHA-256 helps summarize the signing recipe.",
            "ECDSA uses the secret key to make a signature.",
            "Only the matching public key can verify it successfully.",
        ],
        "callout": "The private key does not need to be sent with the message.",
        "badge": "SIGN",
    },
    {
        "title": "Shape the signature bytes",
        "subtitle": "The crypto library returns two big numbers named r and s.",
        "bullets": [
            "The program decodes the library's DER signature form.",
            "Each number becomes exactly 32 bytes for P-256.",
            "Joining them makes COSE's 64-byte ES256 signature value.",
        ],
        "callout": "Same math, different packaging.",
        "badge": "R + S",
    },
    {
        "title": "Put everything in COSE_Sign1",
        "subtitle": "COSE_Sign1 is a standard single-signer envelope.",
        "bullets": [
            "It carries protected headers and unprotected headers.",
            "It carries the message itself because this payload is attached.",
            "It carries the 64-byte signature value.",
        ],
        "callout": "The certificate goes in CB-AdES uHeaders as x5chain.",
        "badge": "THE ENVELOPE",
    },
    {
        "title": "Tag it and save CBOR",
        "subtitle": "CBOR tag 18 tells readers this object is COSE_Sign1.",
        "bullets": [
            "The object is encoded into compact binary CBOR.",
            "The result is written under generated/.",
            "The output file ends in .cbor.",
        ],
        "callout": "That is why opening the raw file in a text editor looks odd.",
        "badge": "SAVE IT",
    },
    {
        "title": "The checker opens the envelope",
        "subtitle": "Validation starts by decoding the signature bytes.",
        "bullets": [
            "It checks the CBOR tag is 18.",
            "It checks the COSE_Sign1 object has four parts.",
            "Unexpected shapes stop validation with a clear error.",
        ],
        "callout": "Never assume input bytes have the shape you expected.",
        "badge": "VALIDATOR",
    },
    {
        "title": "Check the message and signature size",
        "subtitle": "The validator checks basic types before doing crypto.",
        "bullets": [
            "The payload must match the original test document exactly.",
            "The signature must be bytes with a length of 64.",
            "Headers must use the expected CBOR map shapes.",
        ],
        "callout": "Early checks catch broken or surprising input.",
        "badge": "SHAPE CHECK",
    },
    {
        "title": "Check the protected clues",
        "subtitle": "The validator reads the protected header map.",
        "bullets": [
            "It checks that the algorithm is ES256.",
            "It checks that the CWT Claims header contains an integer iat.",
            "It checks the x5t fingerprint has the expected form.",
        ],
        "callout": "A missing required clue means the test signature is rejected.",
        "badge": "HEADER CHECK",
    },
    {
        "title": "Match certificate and fingerprint",
        "subtitle": "The checker follows x5chain and compares the certificate hash.",
        "bullets": [
            "Only the expected uHeaders entry is allowed here.",
            "The embedded DER certificate is loaded.",
            "Its SHA-256 must match the protected x5t value.",
        ],
        "callout": "This links the public certificate to the signed headers.",
        "badge": "CERTIFICATE CHECK",
    },
    {
        "title": "Check the key and verify the math",
        "subtitle": "The public key in the certificate checks the COSE signature.",
        "bullets": [
            "The key must be an elliptic-curve P-256 public key.",
            "The signature is converted back to the library's DER form.",
            "The certificate's self-signature is checked as well.",
        ],
        "callout": "Passing math checks does not make the certificate trusted.",
        "badge": "VERIFY",
    },
    {
        "title": "When something goes wrong",
        "subtitle": "The program reports expected input and decoding problems instead of pretending all is well.",
        "bullets": [
            "Missing files and file-writing problems are reported.",
            "Unexpected CBOR shapes or invalid values stop the run.",
            "The command exits with an error code when creation or checking fails.",
        ],
        "callout": "A clear stop is safer than a pretend success.",
        "badge": "ERROR HANDLING",
    },
    {
        "title": "Tests try a little mischief",
        "subtitle": "The unit test checks both a good signature and tampering.",
        "bullets": [
            "The original payload should verify.",
            "A changed message should fail.",
            "A flipped signature byte should fail too.",
        ],
        "callout": "That is tamper detection: changes are caught, not repaired.",
        "badge": "TESTS",
    },
    {
        "title": "Run it, and remember the limits",
        "subtitle": "You have seen the whole journey from message to checked signature.",
        "bullets": [
            "Install dependencies, then run tools/create_test_signature.py.",
            "Output: generated/CB-AdES-Signature-Test.cbor.",
            "This is a self-signed local demo—not ETSI-approved or PKI-trusted.",
        ],
        "callout": "No trusted timestamp or revocation data is included.",
        "badge": "THE END!",
    },
]


def hex_color(value: str) -> RGBColor:
    value = value.lstrip("#")
    return RGBColor(*(int(value[index : index + 2], 16) for index in (0, 2, 4)))


def add_text(slide, text, x, y, width, height, size, color, bold=False, align=None):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(width), Inches(height))
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(0.02)
    frame.margin_right = Inches(0.02)
    frame.margin_top = Inches(0.02)
    frame.margin_bottom = Inches(0.02)
    paragraph = frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.name = "Aptos"
    paragraph.font.size = Pt(size)
    paragraph.font.bold = bold
    paragraph.font.color.rgb = hex_color(color)
    if align is not None:
        paragraph.alignment = align
    return shape


def add_round_rect(slide, x, y, width, height, fill, radius=True):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE
    shape = slide.shapes.add_shape(
        shape_type, Inches(x), Inches(y), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_color(fill)
    shape.line.fill.background()
    return shape


def build_powerpoint() -> None:
    presentation = Presentation()
    presentation.slide_width = Inches(13.333)
    presentation.slide_height = Inches(7.5)
    blank = presentation.slide_layouts[6]

    for index, item in enumerate(SLIDES):
        slide = presentation.slides.add_slide(blank)
        accent = PALETTE[index % len(PALETTE)]
        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = hex_color(PAPER)

        add_round_rect(slide, 0, 0, 13.333, 0.14, accent, radius=False)
        add_text(slide, f"CODE ADVENTURE  •  {index + 1:02d}/30", 0.65, 6.99, 5, 0.28, 11, MUTED, True)
        add_text(slide, item["title"], 0.65, 6.10, 12.0, 0.62, 28, NAVY, True)
        add_text(slide, item["subtitle"], 0.69, 5.68, 11.8, 0.36, 15, MUTED)

        # A friendly workshop card with a simple robot mascot.
        add_round_rect(slide, 0.65, 1.17, 3.55, 4.12, WHITE)
        add_round_rect(slide, 0.92, 4.53, 3.00, 0.52, accent)
        add_text(slide, item["badge"], 1.02, 4.66, 2.80, 0.23, 12, WHITE, True, PP_ALIGN.CENTER)
        circle = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(1.48), Inches(2.48), Inches(1.88), Inches(1.66)
        )
        circle.fill.solid()
        circle.fill.fore_color.rgb = hex_color(accent)
        circle.line.fill.background()
        add_round_rect(slide, 1.86, 2.05, 1.11, 0.95, NAVY)
        for eye_x in (2.12, 2.55):
            eye = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(eye_x), Inches(2.38), Inches(0.15), Inches(0.15)
            )
            eye.fill.solid()
            eye.fill.fore_color.rgb = hex_color(WHITE)
            eye.line.fill.background()
        add_round_rect(slide, 2.16, 2.70, 0.48, 0.08, ORANGE, radius=False)
        antenna = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(2.37), Inches(1.82), Inches(0.08), Inches(0.25)
        )
        antenna.fill.solid()
        antenna.fill.fore_color.rgb = hex_color(NAVY)
        antenna.line.fill.background()
        spark = slide.shapes.add_shape(
            MSO_SHAPE.OVAL, Inches(2.29), Inches(1.66), Inches(0.24), Inches(0.24)
        )
        spark.fill.solid()
        spark.fill.fore_color.rgb = hex_color(ORANGE)
        spark.line.fill.background()
        add_text(slide, "Byte the robot says:", 1.02, 1.50, 2.9, 0.28, 13, MUTED, True, PP_ALIGN.CENTER)
        add_text(slide, "Let's explore!", 1.10, 3.65, 2.7, 0.36, 20, NAVY, True, PP_ALIGN.CENTER)

        add_round_rect(slide, 4.55, 1.17, 8.12, 4.12, WHITE)
        top = 4.83
        for bullet_index, text in enumerate(item["bullets"]):
            y = top - bullet_index * 1.19
            dot = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(4.94), Inches(y - 0.12), Inches(0.26), Inches(0.26)
            )
            dot.fill.solid()
            dot.fill.fore_color.rgb = hex_color(accent)
            dot.line.fill.background()
            add_text(slide, text, 5.38, y - 0.24, 6.90, 0.70, 18, INK)

        add_round_rect(slide, 0.65, 0.45, 2.38, 0.43, accent)
        add_text(slide, item["badge"], 0.82, 0.56, 2.02, 0.19, 11, WHITE, True, PP_ALIGN.CENTER)
        add_round_rect(slide, 4.55, 0.45, 8.12, 0.45, "#E8EEF8")
        add_text(slide, item["callout"], 4.76, 0.57, 7.72, 0.18, 12, NAVY, True, PP_ALIGN.CENTER)

    presentation.save(POWERPOINT)


def split_text(text: str, font: str, size: float, width: float) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = f"{current} {word}".strip()
        if current and stringWidth(candidate, font, size) > width:
            lines.append(current)
            current = word
        else:
            current = candidate
    if current:
        lines.append(current)
    return lines


def draw_round_rect(pdf, x, y, width, height, fill, radius=18):
    pdf.setFillColor(colors.HexColor(fill))
    pdf.roundRect(x, y, width, height, radius, fill=1, stroke=0)


def draw_robot(pdf, center_x, center_y, accent, scale=1.0):
    pdf.setFillColor(colors.HexColor(NAVY))
    pdf.roundRect(center_x - 38 * scale, center_y - 31 * scale, 76 * scale, 62 * scale, 12 * scale, fill=1, stroke=0)
    pdf.setFillColor(colors.HexColor(WHITE))
    pdf.circle(center_x - 14 * scale, center_y + 5 * scale, 5 * scale, fill=1, stroke=0)
    pdf.circle(center_x + 14 * scale, center_y + 5 * scale, 5 * scale, fill=1, stroke=0)
    pdf.setFillColor(colors.HexColor(ORANGE))
    pdf.roundRect(center_x - 12 * scale, center_y - 13 * scale, 24 * scale, 5 * scale, 2 * scale, fill=1, stroke=0)
    pdf.setStrokeColor(colors.HexColor(NAVY))
    pdf.setLineWidth(4 * scale)
    pdf.line(center_x, center_y + 31 * scale, center_x, center_y + 46 * scale)
    pdf.setFillColor(colors.HexColor(accent))
    pdf.circle(center_x, center_y + 51 * scale, 7 * scale, fill=1, stroke=0)


def draw_wrapped(pdf, text, x, y, width, font="Helvetica", size=16, leading=23, color=INK):
    pdf.setFont(font, size)
    pdf.setFillColor(colors.HexColor(color))
    lines = split_text(text, font, size, width)
    for line in lines:
        pdf.drawString(x, y, line)
        y -= leading
    return y


def build_slides_pdf() -> None:
    page_width, page_height = (960, 540)
    pdf = canvas.Canvas(str(SLIDES_PDF), pagesize=(page_width, page_height))
    pdf.setTitle("The Tiny Signature Machine — 30-slide code adventure")
    pdf.setAuthor("CB-AdES test project")

    for index, item in enumerate(SLIDES):
        accent = PALETTE[index % len(PALETTE)]
        pdf.setFillColor(colors.HexColor(PAPER))
        pdf.rect(0, 0, page_width, page_height, fill=1, stroke=0)
        pdf.setFillColor(colors.HexColor(accent))
        pdf.rect(0, page_height - 10, page_width, 10, fill=1, stroke=0)

        draw_round_rect(pdf, 46, 75, 265, 348, WHITE, 22)
        draw_round_rect(pdf, 68, 99, 220, 42, accent, 12)
        pdf.setFillColor(colors.white)
        pdf.setFont("Helvetica-Bold", 11)
        pdf.drawCentredString(178, 114, item["badge"])
        pdf.setFillColor(colors.HexColor(MUTED))
        pdf.setFont("Helvetica-Bold", 12)
        pdf.drawCentredString(178, 373, "BYTE THE ROBOT SAYS:")
        draw_robot(pdf, 178, 270, accent, 1.1)
        pdf.setFillColor(colors.HexColor(NAVY))
        pdf.setFont("Helvetica-Bold", 18)
        pdf.drawCentredString(178, 192, "Let's explore!")

        pdf.setFillColor(colors.HexColor(NAVY))
        pdf.setFont("Helvetica-Bold", 25 if len(item["title"]) < 30 else 22)
        title_lines = split_text(item["title"], "Helvetica-Bold", 25, 860)
        title_y = 486
        for line in title_lines:
            pdf.drawString(48, title_y, line)
            title_y -= 29
        draw_wrapped(pdf, item["subtitle"], 50, title_y - 1, 850, size=13, leading=17, color=MUTED)

        draw_round_rect(pdf, 338, 75, 576, 348, WHITE, 22)
        bullet_y = 365
        for bullet_index, text in enumerate(item["bullets"]):
            pdf.setFillColor(colors.HexColor(accent))
            pdf.circle(370, bullet_y + 4, 7, fill=1, stroke=0)
            bullet_y = draw_wrapped(
                pdf, text, 393, bullet_y, 488, size=15, leading=21, color=INK
            )
            bullet_y -= 23

        draw_round_rect(pdf, 338, 28, 576, 32, "#E8EEF8", 10)
        pdf.setFillColor(colors.HexColor(NAVY))
        pdf.setFont("Helvetica-Bold", 11)
        callout_lines = split_text(item["callout"], "Helvetica-Bold", 11, 540)
        pdf.drawCentredString(626, 39, callout_lines[0] if callout_lines else "")

        pdf.setFillColor(colors.HexColor(MUTED))
        pdf.setFont("Helvetica", 9)
        pdf.drawString(48, 36, f"CODE ADVENTURE  •  {index + 1:02d}/30")
        pdf.showPage()
    pdf.save()


def draw_guide_page(pdf, page_number: int, title: str, subtitle: str, sections: list[tuple[str, list[str]]], accent: str):
    page_width, page_height = landscape(letter)
    pdf.setFillColor(colors.HexColor(PAPER))
    pdf.rect(0, 0, page_width, page_height, fill=1, stroke=0)
    pdf.setFillColor(colors.HexColor(accent))
    pdf.rect(0, page_height - 11, page_width, 11, fill=1, stroke=0)
    pdf.setFillColor(colors.HexColor(NAVY))
    pdf.setFont("Helvetica-Bold", 25)
    pdf.drawString(42, page_height - 62, title)
    draw_wrapped(pdf, subtitle, 44, page_height - 91, page_width - 88, size=12, leading=16, color=MUTED)

    y = page_height - 132
    for heading, bullets in sections:
        draw_round_rect(pdf, 42, y - 12 - (len(bullets) * 29), page_width - 84, 36 + len(bullets) * 29, WHITE, 14)
        pdf.setFillColor(colors.HexColor(accent))
        pdf.setFont("Helvetica-Bold", 15)
        pdf.drawString(60, y, heading)
        y -= 25
        for bullet in bullets:
            pdf.setFillColor(colors.HexColor(accent))
            pdf.circle(66, y + 4, 3.4, fill=1, stroke=0)
            y = draw_wrapped(pdf, bullet, 80, y, page_width - 150, size=11, leading=15, color=INK)
            y -= 7
        y -= 22

    pdf.setFillColor(colors.HexColor(MUTED))
    pdf.setFont("Helvetica", 9)
    pdf.drawString(42, 24, "A kid-friendly guide to the repository's local CB-AdES test demo.")
    pdf.drawRightString(page_width - 42, 24, f"{page_number} / 4")
    pdf.showPage()


def build_guide_pdf() -> None:
    pdf = canvas.Canvas(str(GUIDE_PDF), pagesize=landscape(letter))
    pdf.setTitle("The Tiny Signature Machine — illustrated guide")
    pdf.setAuthor("CB-AdES test project")
    pages = [
        (
            "The Tiny Signature Machine",
            "A friendly guide to the code that signs one small message and checks it.",
            [
                ("The adventure", [
                    "The sample text is the message, also called the payload.",
                    "The program creates a temporary P-256 key and a self-signed test certificate.",
                    "It builds a COSE_Sign1 envelope, signs it with ES256, and writes a .cbor file.",
                ]),
                ("Where things live", [
                    "test-data/ contains the sample message; tools/ contains Python programs.",
                    "tests/ checks good signatures and makes sure tampering is caught.",
                    "generated/ receives output files and is ignored by Git.",
                ]),
            ],
        ),
        (
            "The key and its test ID",
            "A digital signature uses a secret key to make a seal and a public key to check it.",
            [
                ("Key and certificate", [
                    "The key is generated in memory for each run; the private half is not saved.",
                    "The certificate says Test Signer, Student / Observer and carries the public key.",
                    "It has a short lifetime, a random serial number, and signature-only key usage.",
                ]),
                ("Important trust lesson", [
                    "The certificate signs itself. The program can verify that math.",
                    "A successful math check does not mean a trusted authority vouched for the name.",
                    "This is a local test identity—not ETSI-issued and not PKI-trusted.",
                ]),
            ],
        ),
        (
            "Packing, signing, checking",
            "CBOR is the compact box; COSE describes how the signature travels inside it.",
            [
                ("The signed envelope", [
                    "Protected headers carry ES256, the claimed iat time, and the certificate fingerprint.",
                    "The payload is attached. The certificate is included in CB-AdES uHeaders as x5chain.",
                    "ECDSA signs the COSE Signature1 recipe; the output stores its 64-byte r||s value.",
                ]),
                ("The validator", [
                    "It checks the CBOR tag and COSE_Sign1 shape, then compares the payload.",
                    "It checks header types, the certificate fingerprint, and the P-256 public key.",
                    "It verifies both the message signature and the self-signed certificate math.",
                ]),
            ],
        ),
        (
            "Run it safely",
            "The generated signature is useful for learning and local crypto tests.",
            [
                ("Commands from the repository root", [
                    "Mac/Linux: .venv/bin/python tools/create_test_signature.py",
                    "Windows: .venv\\\\Scripts\\\\python.exe tools\\\\create_test_signature.py",
                    "The signature appears at generated/CB-AdES-Signature-Test.cbor.",
                ]),
                ("What this demo does not provide", [
                    "No trusted TSA timestamp, certificate-chain trust, or revocation evidence.",
                    "The repository does not identify a specific Plugtests portal testcase.",
                    "A cryptographically valid test signature is not proof of real-world trust or ETSI approval.",
                ]),
            ],
        ),
    ]
    for number, (title, subtitle, sections) in enumerate(pages, start=1):
        draw_guide_page(pdf, number, title, subtitle, sections, PALETTE[(number - 1) % len(PALETTE)])
    pdf.save()


def main() -> int:
    try:
        OUTPUT.mkdir(parents=True, exist_ok=True)
        if len(SLIDES) != 30:
            raise ValueError(f"Expected 30 slides, found {len(SLIDES)}.")
        build_powerpoint()
        build_slides_pdf()
        build_guide_pdf()
    except (OSError, ValueError) as exc:
        print(f"Could not create the learning pack: {exc}", file=sys.stderr)
        return 1

    print(f"Created: {GUIDE_PDF}")
    print(f"Created: {SLIDES_PDF}")
    print(f"Created: {POWERPOINT}")
    print("Presentation PDF contains exactly 30 slides.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
