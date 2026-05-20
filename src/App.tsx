import { useState } from 'react'
import Merger from '@/components/merger'
import './App.css'

interface ToolItem {
  id: string
  label: string
  icon: string
  component: React.ComponentType
}

interface ToolCategory {
  label: string
  items: ToolItem[]
}

const TOOL_CATEGORIES: ToolCategory[] = [
  {
    label: '文档工具',
    items: [
      { id: 'excel-word', label: 'Excel→Word 填充', icon: '📋', component: Merger },
    ],
  },
]

const TOOL_MAP = Object.fromEntries(
  TOOL_CATEGORIES.flatMap((c) => c.items.map((t) => [t.id, t.component])),
)

function App() {
  const [activeToolId, setActiveToolId] = useState('excel-word')
  const ActiveTool = TOOL_MAP[activeToolId]

  return (
    <div className="flex h-screen bg-gray-950 overflow-hidden">
      <Sidebar activeId={activeToolId} onSelect={setActiveToolId} />
      <main className="flex-1 overflow-auto">
        {ActiveTool ? <ActiveTool /> : null}
      </main>
    </div>
  )
}

interface SidebarProps {
  activeId: string
  onSelect: (id: string) => void
}

function Sidebar({ activeId, onSelect }: SidebarProps) {
  return (
    <aside className="w-52 bg-gray-900 border-r border-gray-800 flex flex-col flex-shrink-0 select-none">
      {/* Logo */}
      <div className="p-4 border-b border-gray-800">
        <div className="flex items-center gap-2.5">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center text-white text-xs font-bold shadow-lg flex-shrink-0">
            VH
          </div>
          <div>
            <div className="text-sm font-semibold text-gray-100 leading-tight">Virtual Helper</div>
            <div className="text-xs text-gray-500">桌面工具箱</div>
          </div>
        </div>
      </div>

      {/* Navigation */}
      <nav className="p-3 flex-1 overflow-y-auto">
        {TOOL_CATEGORIES.map((category) => (
          <div key={category.label} className="mb-4">
            <div className="text-[10px] text-gray-600 px-2 mb-1 uppercase tracking-widest font-medium">
              {category.label}
            </div>
            {category.items.map((item) => {
              const isActive = item.id === activeId
              return (
                <button
                  key={item.id}
                  onClick={() => onSelect(item.id)}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-lg text-sm transition-colors mb-0.5 ${
                    isActive
                      ? 'bg-indigo-500/15 text-indigo-400 font-medium'
                      : 'text-gray-500 hover:bg-gray-800 hover:text-gray-300'
                  }`}
                >
                  <span className="text-base flex-shrink-0">{item.icon}</span>
                  <span className="truncate text-left">{item.label}</span>
                  {isActive && (
                    <span className="ml-auto w-1.5 h-1.5 rounded-full bg-indigo-400 flex-shrink-0" />
                  )}
                </button>
              )
            })}
          </div>
        ))}
      </nav>

      <div className="p-3 border-t border-gray-800">
        <div className="text-xs text-gray-700 text-center">v1.0.0</div>
      </div>
    </aside>
  )
}

export default App
