def easter_monday(year):
    # It works until 2099
    day = tuple()
    m = 24
    n = 5
    a = year % 19
    b = year % 4
    c = year % 7

    d = (19*a + m) % 30
    e = (2*b + 4*c + 6*d + n) % 7
    if (d + e) < 10:
        day = (99, "Pasquetta", (d + e + 22), "Marzo")
    else:
        day = (99, "Pasquetta", (d + e - 9), "Aprile")
    return day
