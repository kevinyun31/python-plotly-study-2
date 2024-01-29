from http import client
from unittest import result
from pymongo import MongoClient
from bson.son import SON
import pprint

# =======================================================================================

# MongoDB 클라이언트 인스턴스 생성
client = MongoClient(
    host = 'localhost',
    port = 27017,
    username = '',
    password = '' 
)

# 테스트 DB 정의
db = client.test_db

# 데이터 리스트 정의
data = [
    {'x': 1, 'tags': ['dog', 'cat']},
    {'x': 2, 'tags': ['cat']},
    {'x': 2, 'tags': ['mouse', 'cat','dog']},
    {'x': 3, 'tags': []}   
]

# 기존 데이터 초기화
db.things.delete_many({})

# things = 컬렉션에 데이터 벌크 인서트 수행
result = db.things.insert_many(data)

# 추가된 Documents의 ID 확인
print("Inserted Documents ID:", result.inserted_ids)
print()

# =======================================================================================

# 집계 프레임워크

pipeline = [
    {"$unwind": "$tags"},
    {"$group": {"_id": "$tags", "count": {"$sum": 1}}},
    {"$sort": SON([("count", -1), ("_id", -1)])}
]

# 집계 실행 및 결과 출력
agg_result = list(db.things.aggregate(pipeline))
pprint.pprint(agg_result)
print()

# 집계 실행 계획 확인
explain_result = db.command("aggregate", "things", pipeline=pipeline, explain=True)
pprint.pprint(explain_result)

print()
print(" < 모든 작업이 성공적으로 완료되었습니다 > ")
