import { useState, useCallback, useRef, useEffect } from 'react';
import { products, matchMap } from './data/products';
import * as XLSX from 'xlsx';
import './App.css';

const BRAND_NAME = 'öh mo! 어모 · 冰奶·冰乳酪司康集合店';

function App() {
  const [storeName, setStoreName] = useState('');
  const [quantities, setQuantities] = useState({});
  const [aiInput, setAiInput] = useState('');
  const [aiResult, setAiResult] = useState(null); // { matches, unmatched }
  const [showAiDialog, setShowAiDialog] = useState(false);
  const [editableAiResult, setEditableAiResult] = useState([]);
  const aiTextareaRef = useRef(null);

  // 修改数量
  const handleQuantityChange = useCallback((id, value) => {
    const num = parseInt(value) || 0;
    setQuantities(prev => ({ ...prev, [id]: num < 0 ? 0 : num }));
  }, []);

  // 计算总价
  const getTotalPrice = useCallback(() => {
    let total = 0;
    products.forEach(p => {
      const qty = quantities[p.id] || 0;
      total += qty * p.price;
    });
    return total.toFixed(2);
  }, [quantities]);

  // 单项总价
  const getItemTotal = useCallback((id) => {
    const p = products.find(x => x.id === id);
    const qty = quantities[id] || 0;
    return (qty * p.price).toFixed(2);
  }, [quantities]);

  // 有数量的商品
  const getActiveProducts = useCallback(() => {
    return products.filter(p => (quantities[p.id] || 0) > 0);
  }, [quantities]);

  // AI智能识别
  const handleAiRecognize = useCallback(() => {
    const text = aiInput.trim();
    if (!text) return;

    const lines = text.split('\n').filter(l => l.trim());
    const results = [];
    const unmatched = [];

    lines.forEach(line => {
      // 尝试匹配 "数字:商品名 数量 单位" 或 "商品名 数量 单位" 或 "商品名数量单位"
      let found = false;
      const cleaned = line.replace(/^[\d]+[：:、.]?\s*/, '').trim();

      for (const [key, product] of Object.entries(matchMap)) {
        const idx = cleaned.indexOf(key);
        if (idx !== -1) {
          const after = cleaned.substring(idx + key.length).trim();
          const qtyMatch = after.match(/[\d]+/);
          const qty = qtyMatch ? parseInt(qtyMatch[0]) : 1;
          results.push({ name: key, product, qty });
          found = true;
          break;
        }
      }

      // 正向包含匹配
      if (!found) {
        for (const [key, product] of Object.entries(matchMap)) {
          if (cleaned.includes(key)) {
            const after = cleaned.substring(cleaned.indexOf(key) + key.length).trim();
            const qtyMatch = after.match(/[\d]+/);
            const qty = qtyMatch ? parseInt(qtyMatch[0]) : 1;
            results.push({ name: key, product, qty });
            found = true;
            break;
          }
        }
      }

      if (!found) {
        // 尝试数字+商品名模式
        const numMatch = cleaned.match(/^(\d+)\s*/);
        if (numMatch) {
          const num = parseInt(numMatch[1]);
          const rest = cleaned.substring(numMatch[0].length).trim();
          let bestMatch = null;
          let bestLen = 0;
          for (const [key, product] of Object.entries(matchMap)) {
            if (rest.includes(key) && key.length > bestLen) {
              bestMatch = { key, product };
              bestLen = key.length;
            }
          }
          if (bestMatch) {
            results.push({ name: bestMatch.key, product: bestMatch.product, qty: num });
            found = true;
          }
        }
      }

      if (!found) {
        unmatched.push(line);
      }
    });

    if (results.length > 0) {
      setEditableAiResult(results);
      setAiResult({ matches: results, unmatched });
      setShowAiDialog(true);
    } else {
      setAiResult({ matches: [], unmatched });
      setShowAiDialog(true);
      setEditableAiResult([]);
    }
  }, [aiInput]);

  // 确认AI识别结果并填入表格
  const confirmAiResult = useCallback(() => {
    const newQuantities = { ...quantities };
    editableAiResult.forEach(r => {
      const existing = newQuantities[r.product.id] || 0;
      newQuantities[r.product.id] = existing + r.qty;
    });
    setQuantities(newQuantities);
    setShowAiDialog(false);
  }, [editableAiResult, quantities]);

  // 修改AI结果中的数量
  const updateAiQty = useCallback((index, value) => {
    setEditableAiResult(prev => {
      const next = [...prev];
      next[index] = { ...next[index], qty: parseInt(value) || 0 };
      return next;
    });
  }, []);

  // 移除AI结果中的某项
  const removeAiItem = useCallback((index) => {
    setEditableAiResult(prev => prev.filter((_, i) => i !== index));
  }, []);

  // 导出Excel - 采购单
  const exportPurchaseOrder = useCallback(() => {
    const active = getActiveProducts();
    if (active.length === 0) {
      alert('请先填写商品数量');
      return;
    }
    if (!storeName.trim()) {
      alert('请先填写门店名称');
      return;
    }

    const today = new Date();
    const dateStr = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;

    const rows = [
      ['序号', '商品名称', '单价(元)', '单位', '数量', '总价(元)', '备注'],
      ...active.map((p, i) => [
        i + 1,
        p.name,
        p.price,
        p.unit,
        quantities[p.id] || 0,
        (quantities[p.id] || 0) * p.price,
        p.remark || '',
      ]),
      ['', '', '', '', '合计', active.reduce((s, p) => s + (quantities[p.id] || 0) * p.price, 0), ''],
    ];

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { wch: 6 }, { wch: 25 }, { wch: 10 }, { wch: 6 },
      { wch: 8 }, { wch: 12 }, { wch: 20 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '采购单');
    XLSX.writeFile(wb, `${storeName}${dateStr}采购单.xlsx`);
  }, [getActiveProducts, quantities, storeName]);

  // 导出Excel - 数量单
  const exportQuantityOrder = useCallback(() => {
    const active = getActiveProducts();
    if (active.length === 0) {
      alert('请先填写商品数量');
      return;
    }
    if (!storeName.trim()) {
      alert('请先填写门店名称');
      return;
    }

    const today = new Date();
    const dateStr = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;

    const rows = [
      ['序号', '商品名称', '数量', '单位'],
      ...active.map((p, i) => [
        i + 1,
        p.name,
        quantities[p.id] || 0,
        p.unit,
      ]),
    ];

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { wch: 6 }, { wch: 25 }, { wch: 8 }, { wch: 6 },
    ];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '数量单');
    XLSX.writeFile(wb, `${storeName}${dateStr}数量单.xlsx`);
  }, [getActiveProducts, quantities, storeName]);

  // 重置
  const handleReset = useCallback(() => {
    setQuantities({});
    setStoreName('');
    setAiInput('');
  }, []);

  return (
    <div className="app">
      {/* 导航栏 */}
      <nav className="navbar">
        <div className="navbar-brand">
          <span className="navbar-logo">🧊</span>
          <h1 className="navbar-title">{BRAND_NAME}</h1>
        </div>
      </nav>

      <main className="main-content">
        {/* 门店信息 */}
        <section className="card">
          <div className="card-header">
            <span className="card-icon">🏪</span>
            <span>门店信息</span>
          </div>
          <div className="card-body">
            <label className="label">门店名称 <span className="required">*</span></label>
            <input
              type="text"
              className="input"
              placeholder="请输入门店名称"
              value={storeName}
              onChange={e => setStoreName(e.target.value)}
            />
          </div>
        </section>

        {/* AI智能识别 */}
        <section className="card">
          <div className="card-header">
            <span className="card-icon">🤖</span>
            <span>智能识别填表</span>
            <span className="card-hint">（粘贴文字，自动识别商品和数量）</span>
          </div>
          <div className="card-body">
            <textarea
              ref={aiTextareaRef}
              className="textarea"
              rows={4}
              placeholder={`例如:\n1:牛奶40箱\n2:咖奶10箱\n3:茶叶2箱\n4:单杯底座1箱\n5:白桃上允2箱\n6:玫瑰上允10箱\n7:草莓上允10箱\n8:石榴汁1箱`}
              value={aiInput}
              onChange={e => setAiInput(e.target.value)}
            />
            <button className="btn btn-primary btn-full" onClick={handleAiRecognize}>
              🔍 识别并填表
            </button>
          </div>
        </section>

        {/* 商品清单 */}
        <section className="card">
          <div className="card-header">
            <span className="card-icon">📋</span>
            <span>商品清单</span>
            <span className="card-hint">（数量留空默认为0，仅导出数量大于0的商品）</span>
          </div>
          <div className="table-wrapper">
            <table className="product-table">
              <thead>
                <tr>
                  <th className="col-seq">序号</th>
                  <th className="col-name">商品名称</th>
                  <th className="col-price">单价(元)</th>
                  <th className="col-unit">单位</th>
                  <th className="col-qty">数量</th>
                  <th className="col-total">总价(元)</th>
                  <th className="col-remark">备注</th>
                </tr>
              </thead>
              <tbody>
                {products.map(p => {
                  const qty = quantities[p.id] || 0;
                  const itemTotal = (qty * p.price).toFixed(2);
                  const hasQty = qty > 0;
                  return (
                    <tr key={p.id} className={hasQty ? 'row-active' : ''}>
                      <td className="col-seq">{p.id}</td>
                      <td className="col-name">{p.name}</td>
                      <td className="col-price">{p.price.toFixed(2)}</td>
                      <td className="col-unit">{p.unit}</td>
                      <td className="col-qty">
                        <input
                          type="number"
                          className="qty-input"
                          min="0"
                          value={qty || ''}
                          onChange={e => handleQuantityChange(p.id, e.target.value)}
                          placeholder="0"
                        />
                      </td>
                      <td className={`col-total ${hasQty ? 'total-highlight' : 'total-empty'}`}>
                        {hasQty ? itemTotal : '—'}
                      </td>
                      <td className="col-remark">{p.remark}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <div className="card-footer">
            <span className="total-label">总采购金额：</span>
            <span className="total-value">¥{getTotalPrice()}</span>
          </div>
        </section>
      </main>

      {/* 底部固定操作栏 */}
      <div className="bottom-bar">
        <div className="bottom-bar-total">
          <span className="bottom-bar-label">总采购金额</span>
          <span className="bottom-bar-value">¥{getTotalPrice()}</span>
        </div>
        <div className="bottom-bar-actions">
          <button className="btn btn-primary btn-lg" onClick={exportPurchaseOrder}>
            📋 生成采购单
          </button>
          <button className="btn btn-secondary btn-lg" onClick={exportQuantityOrder}>
            📦 导出数量单
          </button>
          <button className="btn btn-outline btn-lg" onClick={handleReset}>
            🔄 重置
          </button>
        </div>
      </div>

      {/* 底部信息 */}
      <footer className="footer">
        {BRAND_NAME} · 采购单生成工具
      </footer>

      {/* AI识别结果弹窗 */}
      {showAiDialog && (
        <div className="modal-overlay" onClick={() => setShowAiDialog(false)}>
          <div className="modal" onClick={e => e.stopPropagation()}>
            <div className="modal-header">
              <h2>AI 识别结果</h2>
              <button className="btn-close" onClick={() => setShowAiDialog(false)}>✕</button>
            </div>
            <div className="modal-body">
              {editableAiResult.length > 0 ? (
                <>
                  <p className="modal-hint">请确认识别结果，可修改数量或删除项目后确认填表</p>
                  <table className="ai-table">
                    <thead>
                      <tr>
                        <th>商品名称</th>
                        <th>数量</th>
                        <th>单价</th>
                        <th>操作</th>
                      </tr>
                    </thead>
                    <tbody>
                      {editableAiResult.map((r, i) => (
                        <tr key={i}>
                          <td>{r.product.name}</td>
                          <td>
                            <input
                              type="number"
                              className="ai-qty-input"
                              min="1"
                              value={r.qty}
                              onChange={e => updateAiQty(i, e.target.value)}
                            />
                          </td>
                          <td>¥{r.product.price.toFixed(2)}/{r.product.unit}</td>
                          <td>
                            <button className="btn btn-sm btn-danger" onClick={() => removeAiItem(i)}>
                              删除
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {aiResult?.unmatched?.length > 0 && (
                    <div className="unmatched-box">
                      <p>⚠️ 以下内容未能识别：</p>
                      <ul>
                        {aiResult.unmatched.map((u, i) => <li key={i}>{u}</li>)}
                      </ul>
                    </div>
                  )}
                </>
              ) : (
                <div className="empty-result">
                  <p>😕 未能识别任何商品，请检查输入格式</p>
                </div>
              )}
            </div>
            {editableAiResult.length > 0 && (
              <div className="modal-footer">
                <button className="btn btn-outline" onClick={() => setShowAiDialog(false)}>取消</button>
                <button className="btn btn-primary" onClick={confirmAiResult}>确认并填表</button>
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
