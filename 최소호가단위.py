###최소 틱 ##

def get_tick_size(price):
    tick_sizes = {
        2000: 1,
        5000: 5,
        10000: 10,
        20000: 10,
        50000: 50,
        100000: 100,
        200000: 100,
    }
    for limit in sorted(tick_sizes.keys()):
        if price < limit:
            return tick_sizes[limit]
    return 500

