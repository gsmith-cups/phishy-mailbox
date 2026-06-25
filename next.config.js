/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  // Standalone output is used for the Docker image. The e2e tests build + run via
  // `next start`, which is incompatible with standalone, so they opt out via this flag.
  // output: process.env.NEXT_DISABLE_STANDALONE === '1' ? undefined : 'standalone',
  // updated code v2:
  output: process.env.NEXT_DISABLE_STANDALONE === '1' ? undefined : 'standalone',
  serverExternalPackages: ['isomorphic-dompurify'],
  experimental: {
    esmExternals: 'loose',
  },
};

module.exports = nextConfig;
