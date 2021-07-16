from flamingo.units import FloatToFeetInchString
from flamingo.units import FeetInchToFloat
from math import pi

tests = [
    ["34", 34],
    ["4.5'", 4.5],
    ["-5", -5],
    ["42'", 42],
    ["- 42'", -42],
    ["15 11", (15.0 + (11.0/12.0))],
    ["0-9", (9.0/12.0)],
    ["7 5 31/64", (7.0 + ((5.0 + (31.0/64.0))/12.0))],
    ["-7 5 31/64", (7.0 + ((5.0 + (31.0/64.0))/12.0)) * -1],
    ["-7' 5.125", (7.0 + (5.125/12.0)) * -1],
    ["5\"", (5.0/12.0)],
    ["-5\"", (5.0/12.0) * -1],
    ["32' 11\"", 32.0 + (11.0/12.0)],
    ["-32' 11\"", (32.0 + (11.0/12.0)) * -1],
    ["0' 6 1/2\"", (6.5 / 12.0)],
    ["- 0' 6 1/2\"", (6.5 / 12.0) * -1],
    ["0-.5", (0.5 / 12.0)],
    ["-0-.5", (0.5 / 12.0) * -1],
    ["0 1/8", (0.125 / 12.0)],
    ["3.14159265359'", pi],
    ["3.14159265359\"", pi / 12.0],
    ["150mm", 0.492126],
]

for test in tests:
    result = FeetInchToFloat(test[0])
    if result:
        testResult = "PASS" if round(result, 5) == round(test[1], 5) else "FAIL"
        print(
            "{}: {} = {} ({}) => {}".format(
                testResult,
                test[0],
                result,
                test[1],
                FloatToFeetInchString(test[1]),
            )
        )
    else:
        print("FAIL: {} was not matched => {}".format(
            test[0],
            FloatToFeetInchString(test[1])
        ))