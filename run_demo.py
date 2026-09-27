from processor import process_file

test_bytes = (
    b"supplier,product,price\n"
    b"Acme,Widget,9.99\n"
    b"Globex,Gadget,12.50"
)

results = process_file(test_bytes)

assert isinstance(results, list)
assert len(results) == 2
assert results[0]["title"] == "Acme"
assert results[1]["title"] == "Globex"
assert isinstance(results[0]["details"], dict)
assert "due_date" not in results[0]["details"]

print("run_demo passed")
