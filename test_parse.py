"""Quick smoke test for the parsers."""

import os
import sys

from edi_parser import EDIFile
from parser_835 import parse_835
from parser_837 import parse_837
from excel_writer import write_combined_excel

# Sample 835 content
SAMPLE_835 = (
    "ISA*00*          *00*          *ZZ*SENDER         *ZZ*RECEIVER       "
    "*230101*1200*^*00501*000000001*0*P*:~"
    "GS*HP*SENDER*RECEIVER*20230101*1200*1*X*005010X221A1~"
    "ST*835*0001~"
    "BPR*I*1500.00*C*ACH*CCP*01*999999999*DA*123456789*9876543210**01*999999999*DA*987654321*20230115~"
    "TRN*1*ABC123456*1234567890~"
    "DTM*405*20230115~"
    "N1*PR*ACME INSURANCE CO*XV*12345~"
    "N1*PE*DR SMITH MEDICAL GROUP*XX*1234567890~"
    "CLP*CLM001*1*500.00*400.00*50.00*12*PAYERCLM001~"
    "NM1*QC*1*DOE*JOHN****MI*MEM001~"
    "NM1*82*1*SMITH*JAMES****XX*1234567890~"
    "DTM*232*20221215~"
    "DTM*233*20221215~"
    "CAS*CO*45*50.00~"
    "CAS*PR*2*50.00~"
    "SVC*HC:99213*250.00*200.00**1~"
    "DTM*472*20221215~"
    "CAS*CO*45*25.00~"
    "CAS*PR*2*25.00~"
    "SVC*HC:99214*250.00*200.00**1~"
    "DTM*472*20221215~"
    "CAS*CO*45*25.00~"
    "CAS*PR*2*25.00~"
    "CLP*CLM002*1*1000.00*800.00*100.00*12*PAYERCLM002~"
    "NM1*QC*1*SMITH*JANE****MI*MEM002~"
    "NM1*82*1*SMITH*JAMES****XX*1234567890~"
    "DTM*232*20221220~"
    "DTM*233*20221220~"
    "CAS*CO*45*100.00~"
    "CAS*PR*3*100.00~"
    "SVC*HC:99215*500.00*400.00**1~"
    "DTM*472*20221220~"
    "CAS*CO*45*50.00~"
    "CAS*PR*3*50.00~"
    "SVC*HC:36415*500.00*400.00**1~"
    "DTM*472*20221220~"
    "CAS*CO*45*50.00~"
    "CAS*PR*3*50.00~"
    "SE*40*0001~"
    "GE*1*1~"
    "IEA*1*000000001~"
)

# Sample 837P content
SAMPLE_837 = (
    "ISA*00*          *00*          *ZZ*SENDER         *ZZ*RECEIVER       "
    "*230201*0800*^*00501*000000002*0*P*:~"
    "GS*HC*SENDER*RECEIVER*20230201*0800*2*X*005010X222A1~"
    "ST*837*0002*005010X222A1~"
    "BHT*0019*00*BATCH001*20230201*0800*CH~"
    "HL*1**20*1~"
    "NM1*85*2*SMITH MEDICAL GROUP*****XX*1234567890~"
    "N3*123 MAIN STREET~"
    "N4*ANYTOWN*CA*90210~"
    "REF*EI*123456789~"
    "HL*2*1*22*1~"
    "SBR*P*18*GRP001*ACME PLAN*****CI~"
    "NM1*IL*1*DOE*JOHN*M***MI*MEM001~"
    "N3*456 OAK AVE~"
    "N4*SOMEWHERE*CA*90211~"
    "DMG*D8*19800115*M~"
    "NM1*PR*2*ACME INSURANCE CO*****PI*12345~"
    "HL*3*2*23*0~"
    "NM1*QC*1*DOE*JIMMY****MI*MEM001D~"
    "N3*456 OAK AVE~"
    "N4*SOMEWHERE*CA*90211~"
    "DMG*D8*20100520*M~"
    "CLM*PAT001*350.00***11:B:1*Y*A*Y*Y~"
    "DTP*431*D8*20230115~"
    "HI*ABK:J06.9*ABF:R50.9~"
    "NM1*82*1*SMITH*JAMES****XX*1234567890~"
    "SV1*HC:99213:25*150.00*UN*1*11~"
    "DTP*472*D8*20230115~"
    "SV1*HC:87880*100.00*UN*1*11~"
    "DTP*472*D8*20230115~"
    "SV1*HC:99050*100.00*UN*1*11~"
    "DTP*472*D8*20230115~"
    "SE*30*0002~"
    "GE*1*2~"
    "IEA*1*000000002~"
)


def test_835():
    print("=" * 50)
    print("Testing 835 Parser")
    print("=" * 50)

    edi = EDIFile(SAMPLE_835)
    print(f"Element sep: '{edi.element_sep}'")
    print(f"Sub-element sep: '{edi.sub_element_sep}'")
    print(f"Segment term: '{edi.segment_term}'")
    print(f"Transaction type: {edi.get_transaction_type()}")
    print(f"Total segments: {len(edi.segments)}")

    parsed = parse_835(edi)
    print(f"\nTransactions: {len(parsed)}")
    for txn in parsed:
        p = txn["payment"]
        print(f"  Payment: ${p['payment_amount']:.2f} via {p['payment_method']}")
        print(f"  Trace #: {p['trace_number']}")
        print(f"  Payer: {p['payer_name']}")
        print(f"  Payee: {p['payee_name']}")
        print(f"  Claims: {len(txn['claims'])}")
        for c in txn["claims"]:
            print(f"    Claim {c['patient_control_number']}: "
                  f"${c['total_charge']:.2f} charged, "
                  f"${c['payment_amount']:.2f} paid, "
                  f"Status: {c['claim_status']}")
            print(f"      Patient: {c['patient_last_name']}, {c['patient_first_name']}")
            print(f"      Services: {len(c['service_lines'])}")
            for svc in c["service_lines"]:
                print(f"        {svc['procedure_code']}: "
                      f"${svc['charge_amount']:.2f} -> ${svc['payment_amount']:.2f}")
                if svc["adjustments"]:
                    for adj in svc["adjustments"]:
                        print(f"          Adj: {adj['group_code']}-{adj['reason_code']} "
                              f"${adj['amount']:.2f} ({adj['reason_description']})")

    output = "/tmp/test_835_output.xlsx"
    write_combined_excel([("test_835.edi", parsed)], [], output)
    print(f"\nExcel written to: {output}")
    print(f"File size: {os.path.getsize(output):,} bytes")


def test_837():
    print("\n" + "=" * 50)
    print("Testing 837 Parser")
    print("=" * 50)

    edi = EDIFile(SAMPLE_837)
    print(f"Transaction type: {edi.get_transaction_type()}")
    print(f"Total segments: {len(edi.segments)}")

    parsed = parse_837(edi)
    print(f"\nTransactions: {len(parsed)}")
    for txn in parsed:
        bp = txn["billing_provider"]
        print(f"  Billing Provider: {bp['name']} (NPI: {bp['npi']})")
        print(f"  Claims: {len(txn['claims'])}")
        for c in txn["claims"]:
            print(f"    Claim {c['claim_id']}: ${c['total_charge']:.2f}")
            print(f"      Patient: {c['patient_name']}")
            print(f"      Subscriber: {c['subscriber_name']} ({c['subscriber_id']})")
            print(f"      Payer: {c['payer_name']}")
            print(f"      Place of Service: {c['place_of_service']}")
            print(f"      Diagnoses: {[d['code'] for d in c['diagnosis_codes']]}")
            print(f"      Service Lines: {len(c['service_lines'])}")
            for svc in c["service_lines"]:
                print(f"        {svc['procedure_code']}: ${svc['charge_amount']:.2f} "
                      f"x {svc['units']} units")

    output = "/tmp/test_837_output.xlsx"
    write_combined_excel([], [("test_837.edi", parsed)], output)
    print(f"\nExcel written to: {output}")
    print(f"File size: {os.path.getsize(output):,} bytes")


if __name__ == "__main__":
    test_835()
    test_837()
    print("\nAll tests passed!")
