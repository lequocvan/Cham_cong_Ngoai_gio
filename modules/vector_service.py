# modules/vector_service.py
# ----------------------------------------------------------------------
# ChromaDB (Vector Database)
# ----------------------------------------------------------------------
import os
import chromadb
CHROMA_DATA_DIR = os.path.join(os.getcwd(), 'chroma_db_data')
os.makedirs(CHROMA_DATA_DIR, exist_ok=True)

client = chromadb.PersistentClient(path=CHROMA_DATA_DIR)

forum_collection = client.get_or_create_collection(
    name="forum_topics",
    metadata={"hnsw:space": "cosine"}
)

def index_topic_to_chroma(topic_id: int, title: str, content: str, category_id: int, author_name: str = "", status: str = "approved"):
    """Lưu hoặc cập nhật bài viết vào ChromaDB"""
    # Chỉ đánh chỉ mục các bài viết đã duyệt
    if status != 'approved':
        return

    document_text = f"Tiêu đề: {title}\nNội dung: {content}"
    forum_collection.upsert(
        documents=[document_text],
        metadatas=[{
            "category_id": int(category_id),
            "author_name": author_name or "N/A",
            "status": status
        }],
        ids=[f"topic_{topic_id}"]
    )

def remove_topic_from_chroma(topic_id: int):
    """Xóa bài viết khỏi ChromaDB khi bài viết bị xóa hoặc bị gỡ duyệt"""
    try:
        forum_collection.delete(ids=[f"topic_{topic_id}"])
    except Exception:
        pass

def search_similar_topics(query: str, top_k: int = 10):
    """Tìm kiếm vector bài viết tương đồng ngữ nghĩa"""
    if not query or not query.strip():
        return []
    
    results = forum_collection.query(
        query_texts=[query],
        n_results=top_k
    )
    
    topic_ids = []
    if results and results.get('ids') and len(results['ids'][0]) > 0:
        for item_id in results['ids'][0]:
            topic_ids.append(int(item_id.replace('topic_', '')))
            
    return topic_ids
