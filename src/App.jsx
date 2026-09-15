import { useEffect, useMemo, useState } from 'react'
import './App.css'

const STORAGE_KEY = 'todo-app.todos'

function loadTodos() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY)
    return raw ? JSON.parse(raw) : []
  } catch {
    return []
  }
}

const FILTERS = {
  all: () => true,
  active: (todo) => !todo.done,
  completed: (todo) => todo.done,
}

function App() {
  const [todos, setTodos] = useState(loadTodos)
  const [text, setText] = useState('')
  const [filter, setFilter] = useState('all')

  useEffect(() => {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(todos))
  }, [todos])

  const visibleTodos = useMemo(
    () => todos.filter(FILTERS[filter]),
    [todos, filter],
  )
  const activeCount = useMemo(() => todos.filter((t) => !t.done).length, [todos])

  function addTodo(e) {
    e.preventDefault()
    const trimmed = text.trim()
    if (!trimmed) return
    setTodos((prev) => [
      { id: crypto.randomUUID(), text: trimmed, done: false },
      ...prev,
    ])
    setText('')
  }

  function toggleTodo(id) {
    setTodos((prev) =>
      prev.map((t) => (t.id === id ? { ...t, done: !t.done } : t)),
    )
  }

  function deleteTodo(id) {
    setTodos((prev) => prev.filter((t) => t.id !== id))
  }

  function clearCompleted() {
    setTodos((prev) => prev.filter((t) => !t.done))
  }

  return (
    <main className="app">
      <h1>Todo</h1>

      <form className="add-form" onSubmit={addTodo}>
        <input
          type="text"
          value={text}
          onChange={(e) => setText(e.target.value)}
          placeholder="할 일을 입력하세요"
          aria-label="새 할 일"
        />
        <button type="submit">추가</button>
      </form>

      {todos.length > 0 && (
        <>
          <ul className="todo-list">
            {visibleTodos.map((todo) => (
              <li key={todo.id} className={todo.done ? 'done' : ''}>
                <label>
                  <input
                    type="checkbox"
                    checked={todo.done}
                    onChange={() => toggleTodo(todo.id)}
                  />
                  <span>{todo.text}</span>
                </label>
                <button
                  type="button"
                  className="delete-btn"
                  aria-label="삭제"
                  onClick={() => deleteTodo(todo.id)}
                >
                  ✕
                </button>
              </li>
            ))}
            {visibleTodos.length === 0 && (
              <li className="empty">표시할 항목이 없습니다</li>
            )}
          </ul>

          <footer className="footer">
            <span>{activeCount}개 남음</span>
            <div className="filters">
              {Object.keys(FILTERS).map((key) => (
                <button
                  key={key}
                  type="button"
                  className={filter === key ? 'active' : ''}
                  onClick={() => setFilter(key)}
                >
                  {key === 'all' ? '전체' : key === 'active' ? '진행중' : '완료'}
                </button>
              ))}
            </div>
            <button type="button" onClick={clearCompleted}>
              완료 항목 삭제
            </button>
          </footer>
        </>
      )}
    </main>
  )
}

export default App
