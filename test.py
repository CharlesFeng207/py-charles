a = [1, 1, 3, 4, 5, 6, 7, 8, 9, 10]

print {'gallahad': 'the pure', 'robin': 'the brave'}

print 'filter', filter(lambda i:i > 3, a)
print 'map', map(lambda i : i + 1, a)
print 'reduce', reduce(lambda i,j: i + j, a)

    
print 'set', set(a)

a_dict = {i : j for i, j in enumerate(a)} #dict comprehensions
print 'a_dict', a_dict

questions = ['name', 'quest', 'favorite color']
answers = ['lancelot', 'the holy grail', 'blue']

print 'zip(questions, answers)', zip(questions, answers)
for q, a in zip(questions, answers):
    print 'What is your {0}?  It is {1}.'.format(q, a)

print 'zip_dict', dict(zip(questions, answers));


