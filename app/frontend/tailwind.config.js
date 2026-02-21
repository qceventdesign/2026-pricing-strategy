/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      colors: {
        qc: {
          50: '#faf6f3',
          100: '#f5f0eb',
          200: '#e8ddd3',
          300: '#d4c4b3',
          400: '#b89d85',
          500: '#846e60',
          600: '#6b5a4f',
          700: '#564940',
          800: '#463c35',
          900: '#2c2620',
        },
      },
    },
  },
  plugins: [],
}
