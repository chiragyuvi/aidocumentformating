import FileUploadCard from '../components/FileUploadCard';

export default function Home() {
  return (
    <main className="relative min-h-screen bg-[#080810] flex flex-col items-center justify-center p-6 overflow-hidden">
      {/* Background */}
      <div className="absolute inset-0 z-0 pointer-events-none">
        <div className="absolute top-0 left-1/2 -translate-x-1/2 w-[700px] h-[400px] bg-[radial-gradient(ellipse_at_top,rgba(99,102,241,0.12),transparent_70%)]" />
        <div className="absolute bottom-0 left-1/2 -translate-x-1/2 w-[500px] h-[300px] bg-[radial-gradient(ellipse_at_bottom,rgba(139,92,246,0.07),transparent_70%)]" />
        <div
          className="absolute inset-0 opacity-[0.025]"
          style={{
            backgroundImage:
              'linear-gradient(rgba(255,255,255,0.05) 1px, transparent 1px), linear-gradient(90deg, rgba(255,255,255,0.05) 1px, transparent 1px)',
            backgroundSize: '40px 40px',
          }}
        />
      </div>

      {/* Header */}
      <div className="relative z-10 max-w-2xl w-full text-center mb-10">
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full border border-indigo-500/30 bg-indigo-500/10 text-indigo-400 text-xs font-medium tracking-widest uppercase mb-5">
          <span className="w-1.5 h-1.5 rounded-full bg-indigo-400 animate-pulse" />
          AI-Powered
        </div>
        <h1 className="text-5xl font-black tracking-tight text-white leading-none">
          Document<br />
          <span className="text-transparent bg-clip-text bg-gradient-to-r from-indigo-400 to-violet-400">
            Formatter
          </span>
        </h1>
        <p className="mt-4 text-zinc-500 text-sm max-w-sm mx-auto">
          Upload your documents and let AI apply consistent, intelligent formatting in seconds.
        </p>
      </div>

      {/* Card */}
      <div className="relative z-10 w-full flex justify-center">
        <FileUploadCard />
      </div>

      <footer className="relative z-10 mt-10 text-zinc-700 text-xs tracking-wide">
        AI Document Formatter • 2026
      </footer>
    </main>
  );
}