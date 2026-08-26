import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";

const geistSans = Geist({ variable: "--font-geist-sans", subsets: ["latin"] });
const geistMono = Geist_Mono({ variable: "--font-geist-mono", subsets: ["latin"] });

export const dynamic = "force-dynamic";

export function generateMetadata(): Metadata {
  const origin = process.env.SITE_ORIGIN ?? "http://localhost:3000";
  const previewImage = new URL("/og.png", origin).toString();
  const title = "IU Work Tracker";
  const description = "Log it once. Use it everywhere. A focused operating system for IU work, projects, impact, and reporting.";
  return {
    metadataBase: new URL(origin), title, description,
    icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
    openGraph: { title, description, type: "website", url: origin, images: [{ url: previewImage, width: 1731, height: 909, alt: "IU Work Tracker — Log it once. Use it everywhere." }] },
    twitter: { card: "summary_large_image", title, description, images: [previewImage] },
  };
}

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en"><body className={`${geistSans.variable} ${geistMono.variable}`}>{children}</body></html>;
}
