import Sidebar from '@/components/shell/sidebar'
import CharityBootstrap from '@/components/charity/charity-bootstrap'

export default function AppLayout({ children }: { children: React.ReactNode }) {
  return (
    <>
      <CharityBootstrap />
      <div className="min-h-screen bg-white md:flex">
        <Sidebar />
        <main className="flex-1 min-w-0 px-5 py-8 md:px-12 md:py-10">
          <div className="max-w-[1240px]">{children}</div>
        </main>
      </div>
    </>
  )
}
