import pprint
from urllib.parse import quote_from_bytes
from pymongo import MongoClient
import datetime


# 1. DB연결 - MongoClient
# 방법1 - URI
# mongodb_URI = "mongodb://localhost:27017/"
# client = MongoClient(mongodb_URI)

# 방법2 - HOST, PORT
client = MongoClient(host='localhost', port=27017)

print(' /// ', client.list_database_names())
# ['admin', 'config', 'local', 'video']


# 2. DB 접근
# 방법1
# db = client.mydb

# 방법2
db = client['mydb']


# 3. Collection 접근
# 방법1
# collection = db.myCol

# 방법2
collection = db['mycol']


# 4. Documents 생성
post = {'author': 'Mike',
        'text': 'My firtst blog post!',
        'tags': ['mongodb', 'python', 'pymongo'],
        'date': datetime.datetime.utcnow()
        }


# 5. Collection 접근 및 Document 추가
# Collection 접근 - 'posts' collection
posts = db.posts

# Document 추가 - insert_one() 메서드 이용
post_id = posts.insert_one(post).inserted_id
print(f' /// Inserted document ID: {post_id}')

# Collection 리스트 조회
print(' /// Collections in the database:', db.list_collection_names())

# # Collection 내 단일 Document 조회
# import pprint
pprint.pprint(posts.find_one())


# # 쿼리를 통한 Documents 조회
# pprint.pprint(posts.find_one({'author': "Mike"}))

# # _id를 통한 Documents 조회 - _id는 binary json 타입으로 조회해야 함
# pprint.pprint(posts.find_one({'_id': 'post_id'}))
# print(type(post_id))

# # _id 값이 str인 경우 조회 안 됨
# post_id_as_str = str(post_id)
# pprint.pprint(posts.find_one({"_id": post_id_as_str}))

# # _id 값이 str인 경우 bson(binary json) 변환 후 조회
# from bson.objectid import ObjectId
# bson_id = ObjectId(post_id_as_str)
# pprint.pprint(posts.find_one({"_id": bson_id}))

# # collection 내 모든 Documents조회
# for post in posts.find():
#     pprint.pprint(post)

# 쿼리를 통한 Documents 조회
# for post in posts.find({'author': 'Mike'}):
#     pprint.pprint(post)
    
    
# 6. 카운팅
# 컬렉션 내 도큐먼트 수 조회
document_count = posts.count_documents({})
print(f" /// Number of documents in the collection: {document_count}")  

# 쿼리를 통한 도큐먼트 수 조회
query_document_count = posts.count_documents({"author": "Mike"})
print(f" /// Number of documents with author 'Mike': {query_document_count}")

# 7. 범위 쿼리
# MongoDB는 여러 고급 쿼리를 사용할 수 있다.
# 특정 날짜 이전의 Document들을 author별로 정렬
d = datetime.datetime(2024, 1, 24, 23, 23, 59) # (year, month, day, hour, minute, second)
for post in posts.find({'date': {'$lt': d}}).sort('author'):
    pprint.pprint(post)
 

print(" < 모든 작업이 성공적으로 완료되었습니다 > ")