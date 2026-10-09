from decimal import Decimal, localcontext
from .numbers import amount


def validate_data(config, records):
    checks, errors = [], []
    by_id = {r["id"]: r for r in records}
    for field in config["fields"]:
        record = by_id.get(field["id"])
        if record is None:
            continue
        for bound, operator in [("min", lambda a, b: a >= b), ("max", lambda a, b: a <= b)]:
            if bound in field and not operator(record["value"], amount(field[bound])):
                errors.append(f"{field['id']}={record['value']} 不符合 {bound}={field[bound]}")
    with localcontext() as ctx:
        ctx.prec = 60
        for check in config["checks"]:
            missing = [x for x in check["left"] + check["right"] if x not in by_id]
            if missing:
                checks.append({"name": check["name"], "status": "unavailable", "missing": missing})
                errors.append(f"{check['name']} 无法验证，缺少 {missing}")
                continue
            left = sum((by_id[x]["value"] for x in check["left"]), Decimal(0))
            right = sum((by_id[x]["value"] for x in check["right"]), Decimal(0))
            difference = left - right
            passed = abs(difference) <= amount(check["tolerance"])
            checks.append({"name": check["name"], "left": left, "right": right,
                           "difference": difference, "tolerance": check["tolerance"],
                           "status": "passed" if passed else "failed"})
            if not passed:
                errors.append(f"{check['name']} 失败，差额 {difference}（Excel 来源单位）")
    return checks, errors
