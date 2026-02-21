/** @type {import('tailwindcss').Config} */
export default {
  content: ["./index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: {
    extend: {
      colors: {
        brand: {
          50: "#fdf8f4",
          100: "#f5e6d3",
          200: "#e8cba8",
          300: "#d4a574",
          400: "#c08a50",
          500: "#a67035",
          600: "#7a6855",
          700: "#4a3728",
          800: "#3a2a1d",
          900: "#2a1d12",
        },
      },
    },
  },
  plugins: [],
};
