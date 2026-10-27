/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: false,
  transpilePackages: [
    '@univerjs/presets',
    '@univerjs/preset-sheets-core',
    '@univerjs/preset-sheets-filter',
    '@univerjs/preset-sheets-sort',
    '@univerjs/preset-sheets-data-validation',
    '@univerjs/preset-sheets-conditional-formatting',
    '@univerjs/preset-sheets-find-replace',
  ],
};

export default nextConfig;
