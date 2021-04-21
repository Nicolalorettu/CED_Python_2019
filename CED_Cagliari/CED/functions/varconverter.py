def fromquerytodict(crs):
    diz = []
    columns = tuple([d[0] for d in crs.description])
    for row in crs:
        diz.append(dict(zip(columns, row)))
    return diz
