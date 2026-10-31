import type { Metadata } from 'next';
import '@fontsource/eb-garamond/500.css';
import '@fontsource/eb-garamond/600.css';
import '@fontsource/eb-garamond/700.css';
import './globals.css';

export const metadata: Metadata = {
  title: 'Restaurant Menu Master Workbook',
  description:
    'A fully relational, offline-first operational analysis and pricing master workbook for restaurant menu customization, complete with an interactive order pricing simulator, automated rule lookup, data integrity health checks, and POS CSV flat exporting.',
  openGraph: {
    title: 'Restaurant Menu Master Workbook',
    description:
      'A fully relational, offline-first operational analysis and pricing master workbook for restaurant menu customization, complete with an interactive order pricing simulator, automated rule lookup, data integrity health checks, and POS CSV flat exporting.',
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body suppressHydrationWarning>{children}</body>
    </html>
  );
}
