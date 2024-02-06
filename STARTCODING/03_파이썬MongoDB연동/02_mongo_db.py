from multiprocessing.reduction import duplicate
import pprint
from urllib.parse import quote_from_bytes
from pymongo import MongoClient, errors
import datetime

# 1. DB연결 - MongoClient
client = MongoClient(host='localhost', port=27017)

print(' /// ', client.list_database_names())
print()

# 2. DB 접근
db = client['mydb']

# 3. Collection 접근 (profiles)
profiles_collection = db['profiles']

# 4. Documents 생성
post = {'author': 'Mike',
        'text': 'My first blog post!',
        'tags': ['mongodb', 'python', 'pymongo'],
        'date': datetime.datetime.utcnow()
        }
print()

# 5. Collection 접근 및 Document 추가 (posts)
posts = db.posts

# Document 추가 - insert_one() 메서드 이용
post_id = posts.insert_one(post).inserted_id
print(f' /// Inserted document ID: {post_id}')

# 6. 카운팅
document_count = posts.count_documents({})
print(f" /// Number of documents in the collection: {document_count}")
print()

# 7. 범위 쿼리
# 특정 날짜 이전의 Document들을 author별로 정렬
d = datetime.datetime(2024, 1, 24, 23, 23, 59)  # (year, month, day, hour, minute, second)
for post in posts.find({'date': {'$lt': d}}).sort('author'):
    pprint.pprint(post)
print()

# 8. 인덱싱
import pymongo

# Indexing
try:
    result = profiles_collection.create_index([('user_id', pymongo.ASCENDING)], unique=True)
    print(" /// Created indexes:")
    for index_name in sorted(profiles_collection.index_information()):
        print(f"   - {index_name}")
except errors.OperationFailure as e:
    print(f"Error creating index: {e}")
print()

# 9. 도큐먼트 초기화 및 추가
# 전체 도큐먼트 삭제
# profiles_collection.delete_many({}) # 위험 전체가 삭제 됨(백업 후 이용권장)
# 이미 존재하는 user_id를 가진 도큐먼트 삭제
profiles_collection.delete_many({'user_id': {'$in': [211, 212, 213]}})

# 중복된 user_id를 가진 도큐먼트가 없을 때만 추가
user_profiles = [
    {'user_id': 211, 'name': 'Luke'},
    {'user_id': 212, 'name': 'Mentis'}
]

for profile in user_profiles:
    existing_profile = profiles_collection.find_one({'user_id': profile['user_id']})
    if existing_profile is None:
        result = profiles_collection.insert_one(profile)
        print(f"Profile added: {profile}")
    else:
        print(f"Profile with user_id {profile['user_id']} already exists.")

# 이미 컬렉션에 있는 user_id 추가 방지
new_profile = {'user_id': 213, 'name': 'Drew'}
duplicate_profile = {'user_id': 212, 'name': 'Tommy'}

# user_id 신규 이므로 정상 추가 됨
result = profiles_collection.insert_one(new_profile)

# 중복된 user_id를 가진 도큐먼트가 없을 때만 추가
for profile in [duplicate_profile]:  # 이미 추가된 new_profile을 제외
    existing_profile = profiles_collection.find_one({'user_id': profile['user_id']})
    if existing_profile is None:
        try:
            result = profiles_collection.insert_one(profile)
            print(f"Profile added: {profile}")
        except errors.DuplicateKeyError:
            print(f"Profile with user_id {profile['user_id']} already exists.")

# 추가된 도큐먼트 출력
for doc in profiles_collection.find():
    pprint.pprint(doc)

print("커밋 테스트")
print(" < 모든 작업이 성공적으로 완료되었습니다 > ")