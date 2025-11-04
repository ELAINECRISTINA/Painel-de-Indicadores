def convert_br_number(value):
    if isinstance(value, str):
        value = value.strip()
        value =  value.replace('.','').replace(',','.')
        return float(value)
    else:
        return float(value)

def list_is_equal(list1, list2) -> bool:
        if len(list1) != len(list2):
            return False
        for i in range(len(list1)):
            if list1[i] != list2[i]:
                return False
        return True