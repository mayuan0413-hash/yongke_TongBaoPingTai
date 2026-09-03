import type { Metadata } from 'next';
import './globals.css';
export const metadata: Metadata = {
 title: '通报平台 · 工作区',
 description: '用多个 Sheet 组织数据，构建可配置的通报工作区。',
};
export default function RootLayout({children}: Readonly<{children: React.ReactNode}>) {
 return <html lang="zh-CN"><body>{children}</body></html>;
}
