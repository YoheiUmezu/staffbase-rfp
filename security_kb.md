# Staffbase Security Knowledge

## データ保護
### 概要
データは転送時および保存時に暗号化される

### 詳細
- TLS 1.2以上
- AES-256（保存時）

### 制約
- 顧客独自鍵管理（BYOK）は非対応（※要確認）

### RFP回答テンプレ
- Yes, data is encrypted in transit using TLS 1.2 or higher and at rest using AES-256.

### NG回答例（防止）
- ❌「最高レベルの暗号化」←曖昧

### 根拠
- Source: Staffbase Security Whitepaper
