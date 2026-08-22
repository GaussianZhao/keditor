'use client';

import { ChangeEvent, ReactNode, useEffect, useRef, useState } from 'react';
import type { AccountInfo } from '@azure/msal-browser';
import {
  CloudDocument,
  connectOneDrive,
  isOneDriveConfigured,
  listDocuments,
  readDocument,
  restoreOneDriveSession,
  saveDocument,
} from './lib/onedrive';

const starter = `# 欢迎使用 Keditor

一个为墨水屏设计、由 OneDrive 云端保存的 Markdown 编辑器。

## 云端自动保存

连接 OneDrive 后，你的修改会在停止输入片刻后自动同步。

- 正文不会自动保存到浏览器
- 可以在电脑、手机和墨水屏之间继续编辑
- 随时导入或导出标准 Markdown 文件

## 编辑与阅读

按 \`⌘ /\`（Mac）或 \`Ctrl /\`（Windows、Linux）切换编辑和阅读模式。

> 少一点界面，多一点文字。`;

type SaveState = 'disconnected' | 'clean' | 'dirty' | 'saving' | 'error';
type ViewMode = 'edit' | 'read';

function statusText(state: SaveState) {
  return {
    disconnected: 'NOT CONNECTED',
    clean: 'SAVED TO CLOUD',
    dirty: 'WAITING TO SAVE',
    saving: 'SAVING…',
    error: 'SYNC FAILED',
  }[state];
}

function Inline({ text }: { text: string }) {
  const parts = text.split(/(`[^`]+`|\*\*[^*]+\*\*|\[[^\]]+\]\([^)]+\))/g);
  return <>{parts.map((part, index) => {
    if (part.startsWith('`') && part.endsWith('`')) return <code key={index}>{part.slice(1, -1)}</code>;
    if (part.startsWith('**') && part.endsWith('**')) return <strong key={index}>{part.slice(2, -2)}</strong>;
    const link = part.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
    if (link) return <a key={index} href={link[2]} target="_blank" rel="noreferrer">{link[1]}</a>;
    return part;
  })}</>;
}

function MarkdownPreview({ content }: { content: string }) {
  const nodes: ReactNode[] = [];
  let inCode = false;
  let code: string[] = [];
  content.split('\n').forEach((line, index) => {
    if (line.startsWith('```')) {
      if (inCode) { nodes.push(<pre key={`code-${index}`}><code>{code.join('\n')}</code></pre>); code = []; }
      inCode = !inCode;
      return;
    }
    if (inCode) { code.push(line); return; }
    if (line.startsWith('### ')) nodes.push(<h3 key={index}><Inline text={line.slice(4)} /></h3>);
    else if (line.startsWith('## ')) nodes.push(<h2 key={index}><Inline text={line.slice(3)} /></h2>);
    else if (line.startsWith('# ')) nodes.push(<h1 key={index}><Inline text={line.slice(2)} /></h1>);
    else if (line.startsWith('> ')) nodes.push(<blockquote key={index}><Inline text={line.slice(2)} /></blockquote>);
    else if (/^[-*] \[[ xX]\] /.test(line)) nodes.push(<div className="task" key={index}><span>{line[3].toLowerCase() === 'x' ? '■' : '□'}</span><Inline text={line.slice(6)} /></div>);
    else if (/^[-*] /.test(line)) nodes.push(<div className="list-item" key={index}><span>•</span><Inline text={line.slice(2)} /></div>);
    else if (/^\d+\. /.test(line)) nodes.push(<div className="list-item" key={index}><span>{line.match(/^\d+\./)?.[0]}</span><Inline text={line.replace(/^\d+\. /, '')} /></div>);
    else if (line.trim() === '---') nodes.push(<hr key={index} />);
    else if (!line.trim()) nodes.push(<div className="blank-line" key={index} />);
    else nodes.push(<p key={index}><Inline text={line} /></p>);
  });
  if (code.length) nodes.push(<pre key="code-final"><code>{code.join('\n')}</code></pre>);
  return <article className="preview">{nodes}</article>;
}

export default function Home() {
  const [account, setAccount] = useState<AccountInfo | null>(null);
  const [documents, setDocuments] = useState<CloudDocument[]>([]);
  const [active, setActive] = useState<CloudDocument | null>(null);
  const [content, setContent] = useState(starter);
  const [saveState, setSaveState] = useState<SaveState>('disconnected');
  const [mode, setMode] = useState<ViewMode>('edit');
  const [message, setMessage] = useState('');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const lastSaved = useRef(starter);

  async function refreshDocuments(preferredId?: string) {
    const next = await listDocuments();
    setDocuments(next);
    if (preferredId) setActive(next.find((doc) => doc.id === preferredId) ?? null);
    return next;
  }

  useEffect(() => {
    restoreOneDriveSession().then(async (restored) => {
      if (!restored) return;
      setAccount(restored);
      const docs = await refreshDocuments();
      setSaveState('clean');
      if (docs[0]) {
        const body = await readDocument(docs[0].id);
        setActive(docs[0]); setContent(body); lastSaved.current = body;
      }
    }).catch((error: Error) => setMessage(error.message));
  }, []);

  useEffect(() => {
    if (!account || !active || content === lastSaved.current) return;
    setSaveState('dirty');
    const timer = window.setTimeout(async () => {
      setSaveState('saving');
      try {
        const saved = await saveDocument(active.name, content, active.eTag);
        lastSaved.current = content;
        setActive(saved);
        setDocuments((items) => items.map((item) => item.id === active.id ? saved : item));
        setSaveState('clean');
      } catch (error) {
        setSaveState('error');
        setMessage(error instanceof Error ? error.message : 'OneDrive 保存失败');
      }
    }, 1200);
    return () => window.clearTimeout(timer);
  }, [content, account, active]);

  useEffect(() => {
    const warn = (event: BeforeUnloadEvent) => {
      if (saveState !== 'dirty' && saveState !== 'saving' && saveState !== 'error') return;
      event.preventDefault();
    };
    window.addEventListener('beforeunload', warn);
    return () => window.removeEventListener('beforeunload', warn);
  }, [saveState]);

  useEffect(() => {
    const toggleView = (event: KeyboardEvent) => {
      if (event.key !== '/' || (!event.metaKey && !event.ctrlKey)) return;
      event.preventDefault();
      setMode((current) => current === 'edit' ? 'read' : 'edit');
    };
    window.addEventListener('keydown', toggleView);
    return () => window.removeEventListener('keydown', toggleView);
  }, []);

  async function handleConnect() {
    if (!isOneDriveConfigured()) {
      setMessage('需要先在 .env.local 中填写 NEXT_PUBLIC_MICROSOFT_CLIENT_ID，再重启网站。');
      return;
    }
    setMessage('');
    try {
      const signedIn = await connectOneDrive();
      setAccount(signedIn);
      let docs = await refreshDocuments();
      if (!docs.length) {
        const created = await saveDocument('欢迎使用Keditor.md', starter);
        docs = await refreshDocuments(created.id);
      }
      const first = docs[0];
      if (first) {
        const body = await readDocument(first.id);
        setActive(first); setContent(body); lastSaved.current = body;
      }
      setSaveState('clean');
    } catch (error) { setMessage(error instanceof Error ? error.message : '连接失败'); }
  }

  async function openDocument(doc: CloudDocument) {
    if (saveState === 'saving' || saveState === 'dirty') { setMessage('请等待当前文档保存完成'); return; }
    try {
      const body = await readDocument(doc.id);
      setActive(doc); setContent(body); lastSaved.current = body; setSaveState('clean'); setSidebarOpen(false);
    } catch (error) { setMessage(error instanceof Error ? error.message : '文档打开失败'); }
  }

  async function newDocument() {
    if (!account) { await handleConnect(); return; }
    const stamp = new Date().toISOString().slice(0, 16).replace('T', '-').replace(':', '');
    try {
      const created = await saveDocument(`未命名-${stamp}.md`, '# 未命名文档\n\n从这里开始写作。');
      await refreshDocuments(created.id);
      setActive(created); setContent('# 未命名文档\n\n从这里开始写作。'); lastSaved.current = '# 未命名文档\n\n从这里开始写作。'; setSaveState('clean');
    } catch (error) { setMessage(error instanceof Error ? error.message : '新建失败'); }
  }

  async function importMarkdown(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!account) { setMessage('请先连接 OneDrive，再导入文档。'); return; }
    try {
      const body = await file.text();
      const saved = await saveDocument(file.name, body);
      await refreshDocuments(saved.id);
      setActive(saved); setContent(body); lastSaved.current = body; setSaveState('clean');
    } catch (error) { setMessage(error instanceof Error ? error.message : '导入失败'); }
    event.target.value = '';
  }

  function exportMarkdown() {
    const blob = new Blob([content], { type: 'text/markdown;charset=utf-8' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob); link.download = active?.name ?? 'Keditor文档.md'; link.click();
    URL.revokeObjectURL(link.href);
  }

  const title = active?.name.replace(/\.md$/i, '') ?? 'Welcome to Keditor';

  return (
    <main className="app-shell">
      <header className="topbar">
        <button className="brand" onClick={() => setSidebarOpen(!sidebarOpen)} aria-label="打开文档列表">Keditor</button>
        <div className="document-title">{title}</div>
        <div className="top-actions">
          <div className="mode-switch" title="Shortcut: Command or Control + Slash" aria-label="Editor view mode">
            <button className={mode === 'edit' ? 'active' : ''} onClick={() => setMode('edit')}>EDIT</button>
            <button className={mode === 'read' ? 'active' : ''} onClick={() => setMode('read')}>READ</button>
          </div>
          <span className={`save-state state-${saveState}`}>{statusText(saveState)}</span>
          <button
            className={`cloud-button${account ? ' connected' : ''}`}
            onClick={handleConnect}
            title={account ? 'OneDrive connected' : 'Connect OneDrive'}
          >
            <span className="cloud-dot" aria-hidden="true" />
            OneDrive
          </button>
        </div>
      </header>

      {message && <div className="notice" role="alert"><span>{message}</span><button onClick={() => setMessage('')}>关闭</button></div>}

      <div className="workspace">
        <aside className={`sidebar ${sidebarOpen ? 'open' : ''}`}>
          <button className="new-document" onClick={newDocument}>＋ 新建云端文档</button>
          <p className="section-label">ONEDRIVE 文档</p>
          <div className="document-list">
            {documents.map((doc) => <button key={doc.id} className={`document-row ${active?.id === doc.id ? 'active' : ''}`} onClick={() => openDocument(doc)}><span className="doc-icon">#</span><span><strong>{doc.name.replace(/\.md$/i, '')}</strong><small>{doc.lastModifiedDateTime ? new Date(doc.lastModifiedDateTime).toLocaleString('zh-CN', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }) : '云端文档'}</small></span></button>)}
            {!documents.length && <div className="empty-cloud"><strong>{account ? '还没有 Markdown 文档' : '连接后显示文档'}</strong><span>文件保存在 OneDrive 的应用专用文件夹中</span></div>}
          </div>
          <div className="sidebar-footer"><label className="file-button">导入 .md<input type="file" accept=".md,.markdown,text/markdown,text/plain" onChange={importMarkdown} /></label><button onClick={exportMarkdown}>导出</button></div>
        </aside>

        <section className="editor-panel">
          {mode === 'edit' ? <textarea className="editor" aria-label="Markdown 编辑器" spellCheck="false" value={content} onChange={(event) => setContent(event.target.value)} /> : <MarkdownPreview content={content} />}
          <footer className="statusbar"><span>{content.replace(/\s/g, '').length} 字 · {content.split(/\s+/).filter(Boolean).length} 词</span><span>{account ? 'OneDrive 云端自动保存' : '当前内容尚未保存'}</span></footer>
        </section>
      </div>
    </main>
  );
}
