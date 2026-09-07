import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: '입금확인 자동화',
  description: '신청세대명단과 입금내역을 자동으로 매칭하는 웹 도구',
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
