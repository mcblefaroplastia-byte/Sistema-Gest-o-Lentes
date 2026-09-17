import React, { useState } from 'react'
import { supabase } from './lib/supabase'

export default function Login() {
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  async function handleLogin(e) {
    e.preventDefault()
    setError('')

    if (!email || !password) {
      setError('Informe o e-mail e a senha.')
      return
    }

    try {
      setLoading(true)

      const { error } = await supabase.auth.signInWithPassword({
        email: email.trim(),
        password,
      })

      if (error) {
        if (error.message === 'Invalid login credentials') {
          setError('E-mail ou senha incorretos.')
        } else {
          setError(error.message)
        }
        return
      }

      // Não precisa redirecionar manualmente.
      // O App.jsx vai detectar a sessão do Supabase.

    } catch (err) {
      console.error(err)
      setError('Não foi possível entrar no sistema.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="login-page">
      <div className="login-container">

        <section className="login-brand-side">
          <div className="login-logo">
            <div className="login-logo-mark">OC</div>

            <div>
              <strong>OftalmoCastro</strong>
              <span>Gestão de Lentes</span>
            </div>
          </div>

          <div className="login-brand-content">
            <span className="login-label">
              CONTROLE CIRÚRGICO
            </span>

            <h1>
              Gestão segura e rastreável das lentes intraoculares.
            </h1>

            <p>
              Estoque, reservas cirúrgicas, movimentações,
              alertas e relatórios em um único sistema.
            </p>
          </div>

          <small>
            Instituto de Olhos OftalmoCastro
          </small>
        </section>

        <section className="login-form-side">

          <form className="login-card" onSubmit={handleLogin}>

            <div className="login-mobile-logo">
              <div className="login-logo-mark">OC</div>

              <div>
                <strong>OftalmoCastro</strong>
                <span>Gestão de Lentes</span>
              </div>
            </div>

            <div className="login-heading">
              <span>ACESSO AO SISTEMA</span>

              <h2>Bem-vindo</h2>

              <p>
                Entre com seu e-mail e senha para continuar.
              </p>
            </div>

            <label className="login-field">
              <span>E-mail</span>

              <input
                type="email"
                placeholder="seuemail@exemplo.com"
                autoComplete="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
              />
            </label>

            <label className="login-field">
              <span>Senha</span>

              <input
                type="password"
                placeholder="Digite sua senha"
                autoComplete="current-password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
              />
            </label>

            {error && (
              <div className="login-error">
                {error}
              </div>
            )}

            <button
              className="login-button"
              type="submit"
              disabled={loading}
            >
              {loading ? 'Entrando...' : 'Entrar no sistema'}
            </button>

            <small className="login-security">
              Acesso restrito aos usuários autorizados pela OftalmoCastro.
            </small>

          </form>

        </section>

      </div>
    </div>
  )
}