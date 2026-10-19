import { useState, useEffect } from 'react';
import { Card, CardContent, CardHeader, CardTitle } from '@/components/ui/card';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer,
} from 'recharts';
import { RefreshCw, Database, Files, ListChecks, Timer, UploadCloud, Download } from 'lucide-react';

// 最新发布包固定下载直链：永远指向 GitHub Release「最新版」的同名资源。
// 每次重打包后，只要用同名资源 wi-workbench-release.zip 发一个新 Release，此链接自动更新。
const RELEASE_DOWNLOAD_URL =
  'https://github.com/DJXD248/wi-workbench-release/releases/latest/download/wi-workbench-release.zip';

interface DocRow {
  series: string;
  file: string;
  file_no: string;
  edition: string;
  product: string;
  processes: number;
  work_hours: number;
  updated: string;
}

interface OutputRow {
  name: string;
  time: string;
  size_kb: number;
}

interface Snapshot {
  schema: number;
  synced_at: string;
  source_pc: string;
  summary: { series: number; docs: number; processes: number; work_hours: number };
  docs: DocRow[];
  recent_output?: OutputRow[];
}

async function fetchSnapshot(): Promise<Snapshot | null> {
  const res = await fetch('/functions/v1/app/snapshot', { credentials: 'same-origin' });
  const text = await res.text();
  if (text.trimStart().startsWith('<')) throw new Error('需要登录站点后才能读取数据');
  if (!res.ok) throw new Error(`服务返回 ${res.status}`);
  const body = JSON.parse(text);
  if (!body || body.ok !== true) throw new Error('服务返回了无效结果');
  if (body.empty) return null;
  return body.snapshot as Snapshot;
}

const fmtHours = (sec: number) =>
  sec >= 3600 ? `${(sec / 3600).toFixed(1)} 小时` :
  sec >= 60 ? `${Math.round(sec / 60)} 分钟` : `${sec} 秒`;

export default function App() {
  const [snapshot, setSnapshot] = useState<Snapshot | null>(null);
  const [state, setState] = useState<'loading' | 'ready' | 'error'>('loading');
  const [error, setError] = useState('');

  const load = () => {
    setState('loading');
    fetchSnapshot()
      .then((s) => { setSnapshot(s); setState('ready'); })
      .catch((e: unknown) => { setError(e instanceof Error ? e.message : String(e)); setState('error'); });
  };
  useEffect(load, []);

  if (state === 'loading') {
    return <div className="flex min-h-screen items-center justify-center text-slate-500">正在读取看板数据…</div>;
  }
  if (state === 'error') {
    return (
      <div className="flex min-h-screen flex-col items-center justify-center gap-3 text-slate-600">
        <p>读取看板数据失败：{error}</p>
        <p className="max-w-md text-center text-sm">若还没同步过数据，请先在作业指导书工作台点「同步进度到云端」。</p>
        <Button onClick={load}><RefreshCw className="mr-1 h-4 w-4" />重试</Button>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-100 p-4 sm:p-8">
      <div className="mx-auto max-w-5xl space-y-6">
        <header className="flex flex-wrap items-end justify-between gap-2">
          <div>
            <h1 className="text-2xl font-bold text-slate-800">作业指导书进度看板</h1>
            <p className="mt-1 text-sm text-slate-500">数据由各电脑上的「作业指导书工作台」一键同步，本页只读展示</p>
          </div>
          <div className="flex items-center gap-2">
            <a href={RELEASE_DOWNLOAD_URL} target="_blank" rel="noreferrer">
              <Button size="sm">
                <Download className="mr-1 h-4 w-4" />下载最新发布包
              </Button>
            </a>
            <Button variant="outline" size="sm" onClick={load}>
              <RefreshCw className="mr-1 h-4 w-4" />刷新
            </Button>
          </div>
        </header>

        {!snapshot ? (
          <Card>
            <CardContent className="flex flex-col items-center gap-3 py-16 text-slate-500">
              <UploadCloud className="h-10 w-10" />
              <p className="font-medium">还没有收到任何同步数据</p>
              <p className="max-w-md text-center text-sm">
                打开作业指导书工作台 →「工具箱」→ 点「同步进度到云端」，成功后刷新本页即可看到内容。
              </p>
              <a href={RELEASE_DOWNLOAD_URL} target="_blank" rel="noreferrer">
                <Button><Download className="mr-1 h-4 w-4" />下载最新发布包</Button>
              </a>
            </CardContent>
          </Card>
        ) : (
          <>
            <div className="flex flex-wrap items-center gap-3 text-sm text-slate-600">
              <Badge variant="secondary">最后同步：{snapshot.synced_at}</Badge>
              <Badge variant="secondary">来源：{snapshot.source_pc || '未知电脑'}</Badge>
            </div>

            <div className="grid grid-cols-2 gap-4 sm:grid-cols-4">
              <Stat icon={<Database className="h-5 w-5" />} label="产品系列" value={String(snapshot.summary.series)} />
              <Stat icon={<Files className="h-5 w-5" />} label="指导书配置" value={String(snapshot.summary.docs)} />
              <Stat icon={<ListChecks className="h-5 w-5" />} label="工序总数" value={String(snapshot.summary.processes)} />
              <Stat icon={<Timer className="h-5 w-5" />} label="工时合计" value={fmtHours(snapshot.summary.work_hours)} />
            </div>

            <div className="grid gap-4 lg:grid-cols-5">
              <Card className="lg:col-span-3">
                <CardHeader><CardTitle>各系列工序数量</CardTitle></CardHeader>
                <CardContent className="h-64">
                  <ResponsiveContainer width="100%" height="100%">
                    <BarChart data={seriesChart(snapshot.docs)}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} />
                      <XAxis dataKey="series" fontSize={12} />
                      <YAxis allowDecimals={false} fontSize={12} />
                      <Tooltip />
                      <Bar dataKey="processes" name="工序数" fill="#3b82f6" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </CardContent>
              </Card>
              <Card className="lg:col-span-2">
                <CardHeader><CardTitle>最近生成的 Excel</CardTitle></CardHeader>
                <CardContent className="max-h-64 space-y-2 overflow-auto">
                  {(snapshot.recent_output ?? []).length === 0 && (
                    <p className="py-8 text-center text-sm text-slate-400">暂无生成记录</p>
                  )}
                  {(snapshot.recent_output ?? []).map((o) => (
                    <div key={o.name} className="flex items-baseline justify-between gap-2 border-b border-slate-100 pb-1 text-sm">
                      <span className="truncate text-slate-700">{o.name}</span>
                      <span className="shrink-0 text-xs text-slate-400">{o.time} · {o.size_kb}KB</span>
                    </div>
                  ))}
                </CardContent>
              </Card>
            </div>

            <Card>
              <CardHeader><CardTitle>配置清单（{snapshot.docs.length} 份）</CardTitle></CardHeader>
              <CardContent className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b text-left text-xs uppercase text-slate-400">
                      <th className="py-2 pr-3">系列</th>
                      <th className="py-2 pr-3">配置文件</th>
                      <th className="py-2 pr-3">文件编号</th>
                      <th className="py-2 pr-3">版本</th>
                      <th className="py-2 pr-3">产品规格</th>
                      <th className="py-2 pr-3 text-right">工序</th>
                      <th className="py-2 pr-3 text-right">工时</th>
                      <th className="py-2">最后修改</th>
                    </tr>
                  </thead>
                  <tbody>
                    {snapshot.docs.map((d) => (
                      <tr key={d.file} className="border-b border-slate-100">
                        <td className="py-2 pr-3 font-medium text-slate-700">{d.series}</td>
                        <td className="py-2 pr-3 text-slate-600">{d.file}</td>
                        <td className="py-2 pr-3 text-slate-600">{d.file_no || '—'}</td>
                        <td className="py-2 pr-3">{d.edition ? <Badge variant="outline">{d.edition}</Badge> : '—'}</td>
                        <td className="max-w-40 truncate py-2 pr-3 text-slate-600" title={d.product}>{d.product || '—'}</td>
                        <td className="py-2 pr-3 text-right">{d.processes}</td>
                        <td className="py-2 pr-3 text-right">{fmtHours(d.work_hours)}</td>
                        <td className="py-2 text-xs text-slate-400">{d.updated}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </CardContent>
            </Card>
          </>
        )}
      </div>
    </div>
  );
}

function Stat({ icon, label, value }: { icon: React.ReactNode; label: string; value: string }) {
  return (
    <Card>
      <CardContent className="flex items-center gap-3 py-4">
        <span className="text-blue-600">{icon}</span>
        <div>
          <p className="text-xs text-slate-500">{label}</p>
          <p className="text-lg font-semibold text-slate-800">{value}</p>
        </div>
      </CardContent>
    </Card>
  );
}

function seriesChart(docs: DocRow[]) {
  const m = new Map<string, number>();
  for (const d of docs) m.set(d.series || '未分组', (m.get(d.series || '未分组') ?? 0) + d.processes);
  return [...m.entries()].map(([series, processes]) => ({ series, processes }));
}
