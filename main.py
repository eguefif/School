import Comment.comment as com
import Diagnostic.diagnostic as diag
# import reports.report as rep


print('             -Welcome-')

print('What would like to do:')
print('''
1. Build report cards commentaries
2. Rundown obsverations
3. Student report
''')

while True:
    a = input('>')
    if int(a) in [1, 2, 3]:
        break
    else:
        print('you have to choose between 1 or 3')

if a == '1':
    print('Select a term:')
    print('''
        1. Term 1
        2. Term 2
        3. Term 3
    ''')

    while True:
        periode = input('>')
        if int(periode) in [1, 2, 3]:
            break
        else:
            print('you have to choose between 1 or 2')
    term = 'P' + periode
    com.comments(term)
if a == '2':
    print('Select a term:')
    print('''
        1. Term 1
        2. Term 2
        3. Term 3
    ''')
    while True:
        periode = input('>')
        if int(periode) in [1, 2, 3]:
            break
        else:
            print('you have to choose between 1 or 2')

    print('''
    1. Global rundowns
    2. French
    3. Maths
    ''')
    while True:
        type = input('>')
        if int(type) in [1, 2, 3]:
            break
        else:
            print('You have to choose 1, 2 or 3')
    term = 'P' + periode
    diag.rundown_observations(term, int(type))

'''
if a == '3':
    print('For whom?')
    name = input('>')

    print(
    1. Periode 1
    2. Periode 2
    3. Periode 3
    )
    while True:
        periode = input('>')
        if int(periode) in [1, 2, 3]:
            break
        else:
            print('You have to choose 1, 2 or 3')
    term = 'P' + periode
    rep.report(name, term)

Todo:
Making a summary for one student with force and weakness.
Put everything per skill with two categories
-proficient
-What's next
Make it possible to choose a specific skill such such qs writing:
'''
