# PS-NCDU

**Analisador de espaço em disco para Windows, em PowerShell, com interface web local e árvore navegável em tempo real.**

[![Version](https://img.shields.io/badge/version-6.28-2c6cb0)](https://github.com/Vietnamix/PS-NCDU)
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://github.com/Vietnamix/PS-NCDU)
[![Plataforma](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](https://github.com/Vietnamix/PS-NCDU)
[![Licença](https://img.shields.io/badge/license-MIT-3fa45b)](License.md)

O PS-NCDU é um script PowerShell **de arquivo único e autônomo** que responde em segundos à pergunta «o que está ocupando este disco?». Executado sem nenhum parâmetro, ele inicia um pequeno servidor web local, abre o navegador e deixa você escolher uma pasta ou um disco para analisar. A árvore é construída **ao vivo** durante a varredura, os tamanhos são preenchidos pasta por pasta, e você navega livremente mesmo antes do fim. Inspirado na ferramenta Unix [`ncdu`](https://dev.yorhel.nl/ncdu), pensado para o ecossistema Windows, sem nenhuma dependência externa.

![Interface do PS-NCDU](PS-NCDU_interface_v6.27b.png)

*PowerShell 5.1+ · Windows · Zero instalação · Arquivo único · 42 idiomas*

Outros idiomas: [Français](README.md) · [English](README_EN.md) · [中文](README_ZH.md) · [हिन्दी](README_HI.md) · [Español](README_ES.md) · [العربية](README_AR.md) · [বাংলা](README_BN.md) · [Русский](README_RU.md) · [اردو](README_UR.md) · [Bahasa Indonesia](README_ID.md) · [Deutsch](README_DE.md) · [日本語](README_JA.md) · [Türkçe](README_TR.md) · [Tiếng Việt](README_VI.md) · [한국어](README_KO.md) · [Italiano](README_IT.md)

---

## Sumário

- [Visão geral](#visão-geral)
- [Funcionalidades](#funcionalidades)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Uso](#uso)
- [A interface](#a-interface)
- [Como funciona a varredura](#como-funciona-a-varredura)
- [Arquivos de trabalho](#arquivos-de-trabalho)
- [Solução de problemas](#solução-de-problemas)
- [Roteiro](#roteiro)
- [Contribuir](#contribuir)
- [Licença](#licença)

---

## Visão geral

Diferente das versões 3.x, que produziam um relatório HTML estático para abrir depois, o PS-NCDU agora é uma **aplicação web local**. O script inicia um servidor HTTP em `127.0.0.1` (porta 8787, com recuo automático para uma porta livre), protegido por um token de sessão, e abre a interface no navegador padrão. Tudo é configurado nessa interface: a pasta a analisar, a profundidade, o filtro de exibição, as exclusões, o idioma.

O servidor permanece na máquina local, não fica exposto na rede, e é encerrado com `Ctrl+C` no console ou pelo botão «Encerrar o servidor» na interface.

![Janela de análise do PS-NCDU](PS-NCDU_scan_form_v6.27b.png)

---

## Funcionalidades

### Varredura e navegação
- **Árvore em tempo real**: a estrutura aparece durante a enumeração, e os tamanhos chegam pasta por pasta à medida que são calculados.
- **Navegação livre durante a varredura**: clique numa pasta para entrar, use a trilha de navegação para subir, sem esperar o fim.
- **Profundidade ajustável ou ilimitada**: pré-carregue alguns níveis para uma exibição fluida, ou a árvore completa. Os tamanhos são sempre exatos, seja qual for a profundidade; os níveis não pré-carregados carregam com um clique.
- **Varredura ilimitada intercalada**: em profundidade ilimitada, cada subárvore é enumerada logo antes de ser medida, de modo que os tamanhos aparecem nos primeiros segundos em vez de esperar o percurso de todo o disco.
- **Interrupção**: um botão «Parar» interrompe a varredura em curso e devolve o controle; iniciar uma nova varredura cancela automaticamente a anterior.
- **Pastas cinza clicáveis**: excluídas, junções, protegidas (ACL) ou não pré-carregadas permanecem visíveis e são varridas sob demanda, com uma fila se já houver uma varredura em andamento.
- **Ordem de processamento alinhada à exibição**: o progresso é preenchido de cima para baixo, sem saltos.

### Leitura dos resultados
- **Ordenação Nome / Tamanho**: por tamanho por padrão (os maiores no topo), suavizada durante a varredura para que as linhas não saltem; um clique alterna para a ordenação por nome.
- **Barras de proporção** e percentuais em relação à pasta atual.
- **Contadores recursivos** de subpastas e arquivos por pasta.
- **Arquivos listados sob demanda** ao abrir uma pasta, ordenados por tamanho, limitados aos 1000 maiores.
- **Ícones por tipo de arquivo**: cerca de 120 extensões comuns (imagens, vídeo, áudio, PDF, documentos de escritório, arquivos compactados, código, executáveis, fontes, imagens de disco, bancos de dados, ebooks, certificados, atalhos) para identificar tipos de relance.
- **Filtro de exibição**: ocultar itens abaixo de 1 MB, 100 MB ou 1 GB para legibilidade, sem alterar a varredura.
- **Pontos de estado coloridos**: varrido, previsto, na fila, em andamento, cinza; uma legenda e uma ajuda integradas explicam cada estado.
- **Tema claro / escuro**.

### Janela de análise
- **Explorador de pastas integrado**: unidades, navegação por clique, pasta superior, «Escolher esta pasta». Sem dependência do seletor nativo do Windows, portanto confiável mesmo em acesso remoto.
- **Acessos rápidos** aos perfis de usuário, **recentes** com remoção individual e limpeza, **unidades** com barra de ocupação.
- **Validação do caminho** ao vivo e no início.
- **Exclusões**: lista das pastas de sistema sempre ignoradas, mais um campo para excluir outras durante uma varredura.
- **Configurações lembradas** (caminho, profundidade, filtro, ordenação, idioma) entre sessões.
- **Enter** para iniciar, **Esc** ou botão de fechar para dispensar a janela quando uma varredura já está exibida.

### Idiomas
- **42 idiomas**, cobrindo mais de 80 % da população mundial: inglês, chinês, hindi, espanhol, francês, árabe, bengali, português, russo, urdu, indonésio, alemão, japonês, coreano, italiano, turco, vietnamita, polonês, neerlandês, ucraniano, romeno, tcheco, grego, sueco, húngaro, persa, tailandês, malaio, filipino, suaíli, tâmil, telugo, marati, guzerate, canarês, malaiala, panjabi, hebraico, hauçá, birmanês, amárico, khmer.
- Detecção automática do idioma do sistema, seletor na janela de análise, escolha lembrada.
- Escrita da direita para a esquerda para árabe, urdu, persa e hebraico.
- As traduções dos 30 idiomas mais recentes são de melhor esforço; a revisão por falantes nativos é bem-vinda, em especial para amárico, khmer, birmanês, hauçá e as línguas indianas.

---

## Pré-requisitos

| Elemento   | Detalhe                                                                                                                             |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| Sistema    | Windows 10 / 11 ou Windows Server                                                                                                   |
| PowerShell | 5.1 (Windows PowerShell) ou 7+ (PowerShell Core)                                                                                    |
| Modo       | **FullLanguage** obrigatório (o servidor web depende de `HttpListener`). Veja [Solução de problemas](#solução-de-problemas) para o modo restrito. |
| Permissões | Leitura nas pastas varridas; alguns caminhos de sistema exigem um console de **administrador**                                      |
| Navegador  | Qualquer navegador recente                                                                                                          |

Nenhum módulo externo é necessário.

---

## Instalação

Clone o repositório ou simplesmente baixe o arquivo `ps-ncdu.ps1`:

```powershell
git clone https://github.com/Vietnamix/PS-NCDU.git
cd PS-NCDU
```

O script está codificado em **UTF-8 com BOM**. Não o salve novamente em outra codificação: o PowerShell 5.1 leria o arquivo como ANSI e quebraria os acentos e os idiomas não latinos da interface.

> **Política de execução**: se o Windows bloquear a execução de scripts, autorize-a para a sessão atual:
>
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```
>
> Este comando não altera nada de forma permanente: vale apenas para a janela do PowerShell aberta.

---

## Uso

**Não há parâmetros de linha de comando**. Basta executar o script:

```powershell
.\ps-ncdu.ps1
```

O script:

1. inicia o servidor local e exibe o endereço no console (por exemplo `http://127.0.0.1:8787/?token=...`);
2. abre esse endereço no navegador padrão;
3. apresenta a janela de análise, onde você escolhe a pasta, a profundidade e a exibição, e então clica em «Analisar».

Se o navegador não abrir sozinho, copie o endereço exibido no console. Para parar: `Ctrl+C` no console, ou o botão «Encerrar o servidor» na interface.

Para varrer caminhos de sistema (`C:\Windows`, a raiz de um disco), inicie o console **como administrador**: caso contrário as pastas protegidas aparecerão em cinza.

---

## A interface

- **Cabeçalho**: trilha de navegação, total da pasta atual com contadores, percentual de progresso, botão de ordenação Nome / Tamanho, tema, «Parar» durante uma varredura, «Nova varredura».
- **Árvore**: uma linha por pasta ou arquivo, com ponto de estado, ícone de tipo, nome, barra de proporção, percentual, contadores e tamanho. As pastas cinza são varridas com um clique; um botão «Varrer pastas cinza (N)» processa todas as da pasta atual.
- **Rodapé**: etapa em curso, caminho realmente lido neste instante, profundidade da varredura, cronômetro, e o painel Legenda / Ajuda.
- **Janela de análise** («Nova varredura»): em duas colunas. À esquerda o destino (caminho, explorador integrado, acessos rápidos, recentes, unidades). À direita as opções (profundidade com controle deslizante e modo ilimitado, filtro de exibição, exclusões). O seletor de idioma fica no cabeçalho desta janela.

---

## Como funciona a varredura

O motor trabalha em etapas. Uma **enumeração** descobre a estrutura e a envia à árvore conforme avança; um **cálculo de tamanhos** percorre então cada subárvore de primeiro nível, na ordem de exibição, subindo os tamanhos parciais a todos os ancestrais uma vez por segundo. Os eventos fluem do servidor para a página via SSE.

Algumas escolhas de projeto que vale conhecer:

- **Uma varredura por vez**, deliberadamente: duas varreduras de disco em paralelo se atrasariam mutuamente. As solicitações extras (pastas cinza) entram na fila e são processadas ao final.
- **Servidor de thread única**: a interrupção é cooperativa. Fechar a conexão (botão «Parar», ou nova varredura) faz falhar a próxima escrita do servidor, que aciona um sinalizador verificado nos laços; a parada efetiva leva até um segundo.
- **Junções e pontos de reanálise** são ignorados para evitar laços e contagens duplas.
- **Unidades de rede**: o espaço delas não é consultado na inicialização, o que evita um travamento quando uma unidade mapeada está inacessível (VPN desligada, por exemplo).
- **Enumeração .NET em fluxo** (`EnumerateFiles` / `EnumerateDirectories`) em vez de `Get-ChildItem`, bem mais rápida no PowerShell 5.1. O teto de desempenho continua sendo o de um script interpretado: ferramentas nativas que leem a MFT do NTFS diretamente são muito mais rápidas, e isso é assumido.

---

## Arquivos de trabalho

| Local                                | Função                                          |
| ------------------------------------ | ----------------------------------------------- |
| `%TEMP%\psncdu\psncdu_debug.log`     | Registro detalhado do servidor e das varreduras |
| `%TEMP%\psncdu\history.txt`          | Histórico de caminhos varridos («Recentes»)      |

Pastas de sistema sempre excluídas da varredura: `C:\Windows\WinSxS`, `C:\Windows\Installer`, `C:\$Recycle.Bin`, `C:\System Volume Information`, `C:\Recovery`, `C:\ProgramData\Microsoft\Windows Defender`, `C:\Windows\SoftwareDistribution`. Você pode adicionar outras, durante uma varredura, na seção Exclusões da janela de análise.

---

## Solução de problemas

| Sintoma                                                       | Causa provável / solução                                                                                                                        |
| ------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------- |
| O script não inicia                                           | Política de execução, veja a nota em [Instalação](#instalação).                                                                                 |
| «O servidor web exige o modo FullLanguage»                    | A sessão está em ConstrainedLanguage (política AppLocker / WDAC). Execute a partir de um console não restrito, ou use a versão 3.2 (relatório HTML estático), que funciona em modo restrito. |
| Acentos ou idiomas não latinos quebrados na interface         | O `.ps1` foi salvo novamente sem o BOM UTF-8. Restaure a codificação original.                                                                   |
| Muitas pastas cinza «protegido»                               | Permissões insuficientes. Reinicie o PowerShell **como administrador**.                                                                         |
| O navegador não abre                                          | Abra manualmente o endereço exibido no console (com o token).                                                                                   |
| Página em branco ou servidor travado na inicialização         | Verifique o registro `%TEMP%\psncdu\psncdu_debug.log`. Uma unidade de rede inacessível podia travar versões antigas; corrigido desde a 5.14.    |
| Nada acontece durante uma varredura ilimitada de um disco     | Corrigido desde a 6.17 (enumeração intercalada). Se persistir, verifique a versão exibida no cabeçalho.                                          |

---

## Roteiro

- [ ] Revisão das traduções dos 30 idiomas recentes por falantes nativos
- [ ] Barra de progresso ancorada no espaço realmente ocupado do disco, em vez de faixas de etapas
- [ ] Indicadores de vazão durante a varredura (arquivos por segundo, MB por segundo)
- [ ] Exportação CSV / JSON dos resultados
- [ ] Comparação de duas varreduras ao longo do tempo

*Sugestões são bem-vindas via issues.*

---

## Contribuir

Contribuições são bem-vindas:

1. Faça um *fork* do repositório.
2. Crie um branch (`git checkout -b feature/minha-funcionalidade`).
3. Preserve a codificação **UTF-8 com BOM** e os *here-strings* do PowerShell intactos.
4. Para traduções, cada idioma é um objeto do dicionário `I18N` dentro do script; compare suas chaves com as de `en` para identificar o que falta.
5. Abra um *pull request* descrevendo claramente a mudança.

Para bugs e ideias, abra uma **issue** informando a versão do Windows, a versão do PowerShell, a versão do PS-NCDU exibida no cabeçalho e, se possível, um trecho do registro `%TEMP%\psncdu\psncdu_debug.log`.

---

## Licença

Distribuído sob a licença **MIT**. Veja o arquivo [`License.md`](License.md).

---

## Autor

**[Eric Guiffault](https://eric.guiffault.com)**

Se este projeto for útil para você, considere deixar uma estrela no GitHub.
