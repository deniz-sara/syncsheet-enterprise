import type { AppProps } from 'next/app';
import Head from 'next/head';
import '@/styles.css';

export default function App({ Component, pageProps }: AppProps) {
  return (
    <>
      <Head>
        <title>Ortak Excel</title>
        <meta
          name="description"
          content="Ortak Excel — Gerçek zamanlı ortak çalışma, bulut senkronizasyonu ve akıllı formül asistanı."
        />
      </Head>
      <Component {...pageProps} />
    </>
  );
}
