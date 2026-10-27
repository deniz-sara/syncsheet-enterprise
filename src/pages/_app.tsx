import type { AppProps } from 'next/app';
import Head from 'next/head';
import '@/styles.css';

export default function App({ Component, pageProps }: AppProps) {
  return (
    <>
      <Head>
        <title>SyncSheet Enterprise — Zero-Upload Ortak Çalışma Excel Platformu</title>
        <meta
          name="description"
          content="Dosya yükleme (DLP) engeline takılmadan tarayıcı içi Excel okuma, 100.000+ satır Canvas performansı ve Figma tarzı gerçek zamanlı ortak çalışma platformu."
        />
      </Head>
      <Component {...pageProps} />
    </>
  );
}
