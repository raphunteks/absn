import type { Metadata } from 'next'
import './globals.css' // <-- BARIS INI SANGAT WAJIB ADA

export const metadata: Metadata = {
  title: 'AXAXYZ Attendance',
  description: 'Enterprise Attendance System',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="id">
      <body>{children}</body>
    </html>
  )
}
