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
  const [collapsed, setCollapsed] = useState(false)
  const ActiveTool = TOOL_MAP[activeToolId]

  return (
    <div className="flex h-screen bg-gray-50 dark:bg-gray-950 overflow-hidden">
      <Sidebar
        activeId={activeToolId}
        collapsed={collapsed}
        onSelect={setActiveToolId}
        onToggle={() => setCollapsed((v) => !v)}
      />
      <main className="flex-1 overflow-auto">
        {ActiveTool ? <ActiveTool /> : null}
      </main>
    </div>
  )
}

interface SidebarProps {
  activeId: string
  collapsed: boolean
  onSelect: (id: string) => void
  onToggle: () => void
}

function Sidebar({ activeId, collapsed, onSelect, onToggle }: SidebarProps) {
  return (
    <aside
      className={`${
        collapsed ? 'w-14' : 'w-52'
      } bg-white dark:bg-gray-900 border-r border-gray-200 dark:border-gray-800 flex flex-col flex-shrink-0 select-none transition-all duration-200 overflow-hidden`}
    >
      {/* Logo + Toggle */}
      <div className="flex items-center border-b border-gray-200 dark:border-gray-800 h-[57px] px-3 gap-2.5">
        {collapsed ? (
          <button
            onClick={onToggle}
            title="展开菜单"
            className="group w-8 h-8 rounded-lg bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center text-white text-xs font-bold shadow-lg flex-shrink-0 hover:from-indigo-400 hover:to-purple-500 transition-all mx-auto"
          >
            <span className="group-hover:hidden">VH</span>
            <svg
              xmlns="http://www.w3.org/2000/svg"
              width="15"
              height="15"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              strokeLinecap="round"
              strokeLinejoin="round"
              className="hidden group-hover:block rotate-180"
            >
              <path d="M11 19l-7-7 7-7" />
              <path d="M19 19l-7-7 7-7" />
            </svg>
          </button>
        ) : (
          <>
            <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-indigo-500 to-purple-600 flex items-center justify-center text-white text-xs font-bold shadow-lg flex-shrink-0">
              VH
            </div>
            <div className="min-w-0 flex-1">
              <div className="text-sm font-semibold text-gray-900 dark:text-gray-100 leading-tight truncate">
                Virtual Helper
              </div>
              <div className="text-xs text-gray-500">桌面工具箱</div>
            </div>
            <button
              onClick={onToggle}
              title="收起菜单"
              className="flex items-center justify-center w-7 h-7 rounded-lg text-gray-400 dark:text-gray-600 hover:bg-gray-100 dark:hover:bg-gray-800 hover:text-gray-600 dark:hover:text-gray-400 transition-colors flex-shrink-0"
            >
              <svg
                xmlns="http://www.w3.org/2000/svg"
                width="15"
                height="15"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
              >
                <path d="M11 19l-7-7 7-7" />
                <path d="M19 19l-7-7 7-7" />
              </svg>
            </button>
          </>
        )}
      </div>

      {/* Navigation */}
      <nav className="p-2 flex-1 overflow-y-auto overflow-x-hidden">
        {TOOL_CATEGORIES.map((category) => (
          <div key={category.label} className="mb-3">
            {!collapsed && (
              <div className="text-[10px] text-gray-400 dark:text-gray-600 px-2 mb-1 uppercase tracking-widest font-medium">
                {category.label}
              </div>
            )}
            {collapsed && <div className="h-2" />}
            {category.items.map((item) => {
              const isActive = item.id === activeId
              return (
                <button
                  key={item.id}
                  onClick={() => onSelect(item.id)}
                  title={collapsed ? item.label : undefined}
                  className={`w-full flex items-center rounded-lg text-sm transition-colors mb-0.5 ${
                    collapsed ? 'justify-center px-0 py-2' : 'gap-2.5 px-3 py-2'
                  } ${
                    isActive
                      ? 'bg-indigo-50 dark:bg-indigo-500/15 text-indigo-600 dark:text-indigo-400 font-medium'
                      : 'text-gray-500 hover:bg-gray-100 dark:hover:bg-gray-800 hover:text-gray-700 dark:hover:text-gray-300'
                  }`}
                >
                  <span className="text-base flex-shrink-0">{item.icon}</span>
                  {!collapsed && (
                    <>
                      <span className="truncate text-left">{item.label}</span>
                      {isActive && (
                        <span className="ml-auto w-1.5 h-1.5 rounded-full bg-indigo-500 dark:bg-indigo-400 flex-shrink-0" />
                      )}
                    </>
                  )}
                </button>
              )
            })}
          </div>
        ))}
      </nav>

      {/* Footer: version */}
      <div className="border-t border-gray-200 dark:border-gray-800 p-2">
        {!collapsed && (
          <div className="text-xs text-gray-400 dark:text-gray-700 text-center py-1">v1.0.0</div>
        )}
      </div>
    </aside>
  )
}

export default App
