def load_records(filename):
    datafile = open(filename, 'r')
    return map(lambda x: x.replace('\n', ''), filter(lambda x: x != '\n', datafile.readlines()))

def get_intersection(current, new):
    intersection = []
    for record in new:
        if record in current:
            intersection.append(record)
    return intersection


if __name__ == "__main__":
    current = load_records("current.data")
    new = load_records("new.data")
    print get_intersection(current, new)