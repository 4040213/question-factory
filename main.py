"""试题批量生成工具 - 独立版入口。"""

import uvicorn
from app import create_app

app = create_app()

if __name__ == "__main__":
    print("=" * 50)
    print("  试题批量生成工具")
    print("  访问: http://localhost:8900/")
    print("=" * 50)
    uvicorn.run(app, host="0.0.0.0", port=8900)
