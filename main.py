import os
from functools import partial
import discord
from discord import guild
from discord.ext import commands
import asyncio
from key import token
import gspread
from google.oauth2.service_account import Credentials
import chat_exporter
import io

token = token.get("TOKEN")


intents = discord.Intents.default()
intents.message_content = True
intents.members = True
intents.guild_messages = True
intents.messages = True
intents.guilds = True

bot = commands.Bot(command_prefix="!", intents=intents)

# COGS



@bot.event
async def on_ready():

  await bot.change_presence(activity=discord.Activity(type=discord.ActivityType.playing, name="Olá! Preparados para o ticket V4?"))
  await bot.tree.sync()
  print ('Bot online!')

@bot.tree.command(name="excluir_canal", description="Exclui o canal de texto onde o comando foi executado.")
async def excluir_canal(interaction: discord.Interaction):
    # Verifica se o autor do comando é um membro com permissão para gerenciar canais
    if interaction.user.guild_permissions.manage_channels:
        await interaction.response.send_message(f"<:tick:1233524716728352830> O canal {interaction.channel.name} está sendo excluído **com sucesso**.", ephemeral=True)
        await asyncio.sleep(10)
        await interaction.channel.delete()
        await interaction.response.send_message(f"O canal {interaction.channel.name} foi excluído com sucesso.", ephemeral=True)
    else:
        await interaction.response.send_message("Você não tem permissão para excluir canais neste servidor.", ephemeral=True)


@bot.tree.command(name="falar", description="Faz a Sophi falar")
async def falar(interaction: discord.Interaction, frase: str):
    await interaction.response.send_message("<:tick:1233524716728352830> Comando enviado **com sucesso**!", ephemeral=True)
    await interaction.channel.send(f"{frase}")

@bot.tree.command(name="ajuda", description="Mostra os comandos do bot")
async def ajuda(interaction: discord.Interaction):
  embed = discord.Embed(
      title="Comandos",
      description="Lista de comandos disponíveis:",
      color=0x00000
  )
  embed.add_field(name="ping", value="Responde com 'Pong!'", inline=False)
  embed.add_field(name="falar", value="Faz a Sophi falar", inline=False)

  user = interaction.user.mention
  content = f"{user}"
  await interaction.response.send_message(content=content, embed=embed)


@bot.command()
async def save(ctx: commands.Context, limit: int = 100, tz_info: str = "America/Sao_Paulo", military_time: bool = True):
    transcript = await chat_exporter.export(
        ctx.channel,
        limit=limit,
        tz_info=tz_info,
        military_time=military_time,
        bot=bot,
    )

    if transcript is None:
        return

    transcript_file = discord.File(
        io.BytesIO(transcript.encode()),
        filename=f"transcript-{ctx.channel.name}.html",
    )

    await ctx.send(file=transcript_file)

# FORMULARIO API

scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

sheet_url = "https://docs.google.com/spreadsheets/d/1st_gKAX3ytjp3jQAPEk1RRIZ19lUhagwBaOOHKNbMoE/edit?pli=1#gid=1305524152"
workbook = client.open_by_url(sheet_url)

sheet = workbook.worksheet("Assinaturas")
sheet2 = workbook.worksheet("Tickets")
sheet3 = workbook.worksheet("Equipe - Tickets Assumidos")


@bot.tree.command(name="procurar_info", description="Retorna todas as informações da linha encontrada.")
async def procurar_info(interaction: discord.Interaction, info: str):
    try:
        celula = sheet.find(f"{info}")
        row_values = sheet.row_values(celula.row)
        if len(row_values) >= 9:
            embed = discord.Embed(
                title=f"API Assinaturas - Nexa",
                color=0x00000,
                url="https://agencynexa.com.br"
                )
            embed.add_field(name="Nome/Nick do cliente", value=row_values[0], inline=False)
            embed.add_field(name="E-mail", value=row_values[1], inline=False)
            embed.add_field(name="IP  da máquina", value=row_values[2], inline=False)
            embed.add_field(name="Início da Assinatura", value=row_values[3], inline=False)
            embed.add_field(name="Fim da Assinatura", value=row_values[4], inline=False)
            embed.add_field(name="Plano", value=row_values[5], inline=False)
            embed.add_field(name="Valor pago:", value=row_values[6], inline=False)
            embed.add_field(name="Extras (Armazenamento)", value=row_values[7], inline=False)
            embed.add_field(name="Status da assinatura", value=row_values[8], inline=False)
            embed.set_image(url="https://media.discordapp.net/attachments/1124704207484751994/1229299165432119316/nexa_cinza.png?ex=662c8a10&is=662b3890&hm=a8d940d79ad01f27e4228705f20c6b2d9b67a4c7f2e75e0be9f74c74f0963468&=&format=webp&quality=lossless&width=1025&height=342")
            embed.set_footer(text="Agency Nexa • © Todos os direitos reservados.", icon_url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
            embed.set_thumbnail(url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
            await interaction.response.send_message(embed=embed)
        else:
            await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Número insuficiente de colunas encontradas.")
    except:
        await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Informação não encontrada na planilha.")

@bot.tree.command(name="proximo_em_branco", description="Retorna o número da próxima linha marcada como 'Linha vazia'.")
async def proximo_em_branco(interaction: discord.Interaction):
    try:
        # Encontra a célula marcada como "Linha vazia"
        for i, row in enumerate(sheet.get_all_values(), start=1):
            for j, cell in enumerate(row, start=1):
                if cell.strip().lower() == "linha vazia":
                    await interaction.response.send_message(f"A próxima linha marcada como 'Linha vazia' é a linha {i}.", ephemeral=True)
                    return  # Retorna após encontrar a célula
        # Se não encontrou uma célula marcada como "Linha vazia"
        await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Não foi possível encontrar uma célula marcada como 'Linha vazia'.")
    except Exception as e:
        await interaction.response.send_message(f"Ocorreu um erro: {e}", ephemeral=True)

@bot.tree.command(name="proximo_em_branco_tickets", description="Retorna o número da próxima linha marcada como 'Linha vazia'.")
async def proximo_em_branco(interaction: discord.Interaction):
    try:
        # Encontra a célula marcada como "Linha vazia"
        for i, row in enumerate(sheet2.get_all_values(), start=1):
            for j, cell in enumerate(row, start=1):
                if cell.strip().lower() == "linha vazia":
                    await interaction.response.send_message(f"A próxima linha marcada como 'Linha vazia' é a linha {i}.", ephemeral=True)
                    return  # Retorna após encontrar a célula
        # Se não encontrou uma célula marcada como "Linha vazia"
        await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Não foi possível encontrar uma célula marcada como 'Linha vazia'.", ephemeral=True)
    except Exception as e:
        await interaction.response.send_message(f"Ocorreu um erro: {e}", ephemeral=True)

@bot.tree.command(name="verificar_linhas_ocupadas", description="Verifica e informa as linhas marcadas como 'linha ocupada'.")
async def verificar_linhas_ocupadas(interaction: discord.Interaction):
    try:
        linhas_ocupadas = []
        
        for i, row in enumerate(sheet2.get_all_values(), start=1):
            for j, cell in enumerate(row, start=1):
                if cell.strip().lower() == "linha ocupada":
                    linhas_ocupadas.append((i, row))
                    break
        
        if linhas_ocupadas:
            mensagem = "Linhas marcadas como 'linha ocupada':\n"
            for linha, valores in linhas_ocupadas:
                mensagem += f"Linha {linha}: {', '.join(valores)}\n"
            await interaction.response.send_message(mensagem)
        else:
            await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Nenhuma linha marcada como 'linha ocupada' encontrada.")
    except Exception as e:
        await interaction.response.send_message(f"Ocorreu um erro: {e}")

import datetime

@bot.tree.command(name="teste", description="Teste de tempo")
async def falar(interaction: discord.Interaction):
  horario_atual = datetime.datetime.now()

  # Obter o horário atual
  horario_atual = datetime.datetime.now()

  # Definir o intervalo de tempo (5 segundos)
  intervalo = datetime.timedelta(seconds=6)

  # Adicionar o intervalo ao horário atual
  horario_futuro = horario_atual + intervalo

  # Obter o timestamp do horário futuro
  timestamp_futuro = horario_futuro.timestamp()

  # Arredondar o timestamp para o segundo mais próximo
  timestamp_arredondado = round(timestamp_futuro)


  # Imprimir o timestamp atual
  await interaction.response.send_message(f"<t:{timestamp_arredondado}:R>")

# SUPER HIPER MEGA TESTE

@bot.tree.command(name="ticket", description="Ticket V4 - Created by Thex")
async def ticket(interaction: discord.Interaction):
    if interaction.user.id == 1024741176554823751:
        await interaction.response.send_message("<:tick:1233524716728352830> Mensagem enviada com **sucesso**!", ephemeral=True)
        embed = discord.Embed(
            title="Painel Tickets - Nexa",
            description="""Sejam **bem-vindos** ao nosso painel de tickets! Caso precise de **suporte** abra um **ticket**.\n\n**Horários de Atendimento** Nexa:\n\n**Segunda á Sexta:** 9:30 - 00:00\n**Sábados e Domingos:** 10:00 - 02:00\n\nCaso haja tickets **abertos** antes, ou depois do tempo determinado, terá que **esperar** que nosso horário esteja aberto para responde-lo (a).""",
            color=0x00000,
            url="https://agencynexa.com.br"
        )
        embed.set_footer(text="Agency Nexa • © Todos os direitos reservados.", icon_url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
        embed.set_thumbnail(url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
        embed.set_image(url="https://media.discordapp.net/attachments/1097341485667586069/1216919465229418566/ticket.png?ex=66022391&is=65efae91&hm=ced252dbe390ae70580f740c530515b2b593aa1235bc6168447c0b805082faa3&=&format=webp&quality=lossless&width=1025&height=193")

    else:
        await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Você não tem permissão para usar este comando.", ephemeral=True)
        print("Alguém está tentando executar sem permissão")

    # Botão de Suporte
    ticket_button = discord.ui.Button(
        style=discord.ButtonStyle.blurple,
        label="Suporte",
        emoji="<:suporte:1233336866489630730>",
        custom_id="ticket"
    )
    view = discord.ui.View()
    view.add_item(ticket_button)

    async def criar_Canal(interaction: discord.Interaction, channel_type: discord.ChannelType):
        try:
            guild = interaction.guild
            category_id = 1113989716208005132
            category = guild.get_channel(category_id)

            view = discord.ui.View()
            
            cargo_equipe = guild.get_role(1095021946850250834)  # Obtenha o objeto de cargo para a equipe

            criado_por = interaction.user.mention
            criado_por_id = interaction.user.id
        except:
            print("Erro")
            
        id_encontrado = False
        for row in sheet2.get_all_values():
            for cell in row:
                if cell == str(interaction.user.id):
                    await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Você **já possui** um ticket aberto. \n \n**OBS:** Você pode ter apenas 1 ticket aberto, caso tenha um de suporte, ele precisa ser fechado para abrir um de reembolso.", ephemeral=True)
                    id_encontrado = True
                    break
            if id_encontrado:
                break

        # Se o ID do usuário não foi encontrado e ainda não preenchemos uma célula vazia, cria um novo ticket e adiciona o ID na próxima linha vazia
        if not id_encontrado:
            linha_vazia_encontrada = False
            for i, row in enumerate(sheet2.get_all_values(), start=1):
                for j, cell in enumerate(row, start=1):
                    if cell.strip().lower() == "linha vazia":
                        sheet2.update_cell(i, 1, str(criado_por_id))
                        nome_global_usuario = interaction.user.name
                        sheet2.update_cell(i, 2, nome_global_usuario)
                        
                        
                        linha_vazia_encontrada = True
                        break
                if linha_vazia_encontrada:
                    break

            await interaction.response.send_message("<a:load:1233323192647417908>  Seu ticket está sendo **criado**, aguarde.", ephemeral=True)

            novo_canal = await guild.create_text_channel(
                name=f"🎟️┃ticket-{interaction.user.name}",
                category=category,
                slowmode_delay=0,
                position=0,
                nsfw=False,
                overwrites={
                    guild.default_role: discord.PermissionOverwrite(view_channel=False),
                    cargo_equipe: discord.PermissionOverwrite(view_channel=True, send_messages=False),
                    interaction.user: discord.PermissionOverwrite(view_channel=True, send_messages=True)
                },
                reason=None
            )

            botao_link = discord.ui.Button(
                style=discord.ButtonStyle.link,
                label="Atalho para o ticket",
                url=f"https://discord.com/channels/{guild.id}/{novo_canal.id}",
                emoji="<:join:1233548663746203679>"
            )

            view.add_item(botao_link)

            await interaction.followup.send("<:tick:1233524716728352830>  Seu ticket foi **criado com sucesso** e você foi **mencionado!** \n \n <:info:1196645186315493436> **OBS:** Sistema de ticket novo e em fase de testes, qualquer erro deve ser reportado.", ephemeral=True, view=view)


            mensagem_embed = discord.Embed(
                title="Painel Tickets - Nexa",
                description=f"Olá **{interaction.user.global_name}**, \n \n <:setabrancaarrumada:1233574306294927430> Por favor, ao fazer uma pergunta, forneça uma **descrição completa** e **detalhada** do seu problema ou dúvida. \n \n <:setabrancaarrumada:1233574306294927430> **Evite mensagens curtas** como 'Oi' ou 'Olá', pois essas podem não fornecer informações suficientes para uma resposta adequada.",
                color=0x000000,
                url="https://agencynexa.com.br"
            )
            mensagem_embed.set_footer(text="Agency Nexa • © Todos os direitos reservados.", icon_url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
            mensagem_embed.set_thumbnail(url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
            
            assumir_ticket = discord.ui.Button(
                style=discord.ButtonStyle.blurple,
                label="Assumir Ticket",
                emoji="<:suporte:1233336866489630730>",
                custom_id="assumir_ticket"
            )

            async def assumindoticket(interaction: discord.Interaction):
                if discord.utils.get(interaction.user.roles, id=1095021946850250834):
                    await interaction.response.send_message("<:tick:1233524716728352830> **Sucesso:** você assumiu o ticket.", ephemeral=True)
                    await interaction.channel.send(f"Seu ticket será atendido por {interaction.user.mention}.")
                    await interaction.channel.set_permissions(interaction.user, send_messages=True)
                    
                    id_usuario = interaction.user.id
                    linha_encontrada = None
                    coluna_encontrada = None

                    # Percorre todas as células da planilha
                    for i, row in enumerate(sheet3.get_all_values(), start=1):
                        for j, cell in enumerate(row, start=1):
                            # Verifica se o valor da célula é igual ao ID do usuário
                            if cell == f"{id_usuario}":
                                # Se encontrou o ID do usuário na planilha, armazena a linha e a coluna e sai dos loops
                                linha_encontrada = i
                                coluna_encontrada = j
                                valor_atual = sheet3.cell(i, 3).value
                                print(valor_atual)
                                try:
                                    valor_atual = int(valor_atual)
                                except ValueError:
                                    print("Erro ao converter para valor inteiro")
                                novo_valor_total = valor_atual + 1
                                sheet3.update_cell(i, 3, str(novo_valor_total))
                                print("Número de tickets assumidos atualizado.")
                                break
                        if linha_encontrada is not None:
                            break

                    # Verifica se o usuário foi encontrado na planilha
                    if linha_encontrada is not None:
                        print(f"O usuário foi encontrado na linha {linha_encontrada} e coluna {coluna_encontrada}.")
                    else:
                        for row_index, row in enumerate(sheet3.get_all_values(), start=1):
                            if "linha vazia" in row:
                                sheet3.update_cell(row_index, 1, str(id_usuario))
                                sheet3.update_cell(row_index, 2, interaction.user.name)
                                sheet3.update_cell(row_index, 3, "1")
                                print("Usuário adicionado à planilha.")
                                break
                else:
                    await interaction.response.send_message("Você não é da **equipe de suporte**, apenas pessoas da equipe podem `Assumir Tickets`.", ephemeral=True)

            assumir_ticket.callback = assumindoticket

            fechar_ticket = discord.ui.Button(
                style=discord.ButtonStyle.red,
                label="Fechar Ticket",
                emoji="<:lixeira:1233523785202667642>",
                custom_id="fechar_ticket"
            )
            view = discord.ui.View()
            view.add_item(fechar_ticket)
            view.add_item(assumir_ticket)
            

                

            async def fechandoticket(interaction: discord.Interaction):
                if discord.utils.get(interaction.user.roles, id=1095021946850250834):
                    
                    tz_info = "America/Sao_Paulo"
                    
                    transcript = await chat_exporter.export(
                        novo_canal,
                        tz_info=tz_info,
                    )

                    transcript_file = discord.File(
                                io.BytesIO(transcript.encode()),
                                filename=f"transcript-{novo_canal.name}.html",
                            )
                    
                    horario_atual = datetime.datetime.now()
                    intervalo = datetime.timedelta(seconds=6)
                    horario_futuro = horario_atual + intervalo
                    timestamp_futuro = horario_futuro.timestamp()
                    timestamp_arredondado = round(timestamp_futuro)
                    await interaction.response.send_message(f"<:tick:1233524716728352830> Entendido, este ticket será fechado <t:{timestamp_arredondado}:R>.")
                    fechado_por = interaction.user.mention
                    fechado_por_id = interaction.user.id
                    await asyncio.sleep(5)
                    await interaction.channel.delete()

                    # Remove o ID do usuário da planilha
                    for i, row in enumerate(sheet2.get_all_values(), start=1):
                        if str(row[0]) == str(criado_por_id):
                            sheet2.update_cell(i, 1, "")  # Limpa a célula com o ID do usuário
                            sheet2.update_cell(i, 2, "")
                            break  # Interrompe o loop após encontrar e atualizar a célula
                        
                    server_id = '1195527779651964978'
                    channel_id = '1232547539962691617'

                    server = bot.get_guild(int(server_id))
                    if server is None:
                        print(f"Servidor não encontrado com o ID {server_id}")
                        return

                    channel = server.get_channel(int(channel_id))
                    if channel is None:
                        print(f"Canal não encontrado com o ID {channel_id}")
                        return

                    log_ticket = discord.Embed(
                        title="Ticket de suporte fechado",
                        color=0x00000,
                        url="https://agencynexa.com.br"
                    )

                    hora_atual = datetime.datetime.now()
                    time_atual = hora_atual.timestamp()
                    time_arredondado = round(time_atual)

                    log_ticket.add_field(name="Criado por:", value=f"{criado_por} `{criado_por_id}`", inline=False)
                    log_ticket.add_field(name="Fechado por:", value=f"{fechado_por} `{fechado_por_id}`", inline=False)
                    log_ticket.add_field(name="<:info:1196645186315493436> Acredita que foi um engano? Seu problema não foi resolvido?", value=f"\n Clique no botão abaixo e crie um novo ticket e entre em contato consco novamente!", inline=False)
                    log_ticket.add_field(name="<:relogio:1233543048051556453> Fechado em:", value=f"<t:{time_arredondado}:F>", inline=False)
                    log_ticket.set_footer(text="Agency Nexa • © Todos os direitos reservados.", icon_url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
                    log_ticket.set_thumbnail(url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
                    button_url = discord.ui.Button(
                        style=discord.ButtonStyle.link,
                        label="Criar novo ticket",
                        url="https://discordapp.com/channels/1094399810158739466/1095014679404875918"
                    )

                    view = discord.ui.View()
                    view.add_item(button_url)

                    await channel.send(embed=log_ticket, view=view)
                    await channel.send(file=transcript_file)

                    enviar = bot.get_user(criado_por_id)
                    await enviar.send(embed=log_ticket, view=view)
                    await enviar.send(file=transcript_file)

                else:
                    await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Você não é da **equipe de suporte**, apenas pessoas da equipe podem `Fechar Tickets`.", ephemeral=True)

            fechar_ticket.callback = fechandoticket

            equipe = guild.get_role(1095021946850250834)

            content = f"{interaction.user.mention} {equipe.mention}"

            # Envio da mensagem embed no canal recém-criado
            mensagem_final = await novo_canal.send(embed=mensagem_embed, content=content, view=view)

            await mensagem_final.pin()

    ticket_button.callback = partial(criar_Canal, channel_type=discord.ChannelType.text)

    reembolso_button = discord.ui.Button(
        style=discord.ButtonStyle.green,
        label="Reembolso ",
        emoji="<:dinheiro:1233336975226830858>",
        custom_id="reembolso"
    )
    view.add_item(reembolso_button)

    async def criar_Canal(interaction: discord.Interaction, channel_type: discord.ChannelType):
        try:
            guild = interaction.guild
            category_id = 1193253630518771794
            category = guild.get_channel(category_id)

            view = discord.ui.View()

            cargo_equipe = guild.get_role(1095021946850250834)  # Obtenha o objeto de cargo para a equipe

            criado_por = interaction.user.mention
            criado_por_id = interaction.user.id
        except:
            print("Erro")

        id_encontrado = False
        for row in sheet2.get_all_values():
            for cell in row:
                if cell == str(interaction.user.id):
                    await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Você **já possui** um ticket aberto. \n \n**OBS:** Você pode ter apenas 1 ticket aberto, caso tenha um de suporte, ele precisa ser fechado para abrir um de reembolso.", ephemeral=True)
                    id_encontrado = True
                break
            if id_encontrado:
                break

        # Se o ID do usuário não foi encontrado e ainda não preenchemos uma célula vazia, cria um novo ticket e adiciona o ID na próxima linha vazia
        if not id_encontrado:
            linha_vazia_encontrada = False
            for i, row in enumerate(sheet2.get_all_values(), start=1):
                for j, cell in enumerate(row, start=1):
                    if cell.strip().lower() == "linha vazia":
                        sheet2.update_cell(i, 1, str(criado_por_id))
                        nome_global_usuario = interaction.user.name
                        sheet2.update_cell(i, 2, nome_global_usuario)
                        linha_vazia_encontrada = True
                        break
                if linha_vazia_encontrada:
                    break

            await interaction.response.send_message("<a:load:1233323192647417908>  Seu ticket está sendo **criado**, aguarde.", ephemeral=True)

            novo_canal = await guild.create_text_channel(
                name=f"💰┃reembolso-{interaction.user.name}",
                category=category,
                slowmode_delay=0,
                position=0,
                nsfw=False,
                overwrites={
                guild.default_role: discord.PermissionOverwrite(view_channel=False),
                    cargo_equipe: discord.PermissionOverwrite(view_channel=True, send_messages=False),
                    interaction.user: discord.PermissionOverwrite(view_channel=True, send_messages=True)
                },
                reason=None
            )

            botao_link = discord.ui.Button(
                style=discord.ButtonStyle.link,
                label="Atalho para o ticket",
                url=f"https://discord.com/channels/{guild.id}/{novo_canal.id}"
            )

            view.add_item(botao_link)

            await interaction.followup.send("<:tick:1233524716728352830>  Seu ticket foi **criado com sucesso** e você foi **mencionado!** \n \n <:info:1196645186315493436> **OBS:** Sistema de ticket novo e em fase de testes, qualquer erro deve ser reportado.", ephemeral=True, view=view)


            mensagem_embed = discord.Embed(
                title="Painel Tickets - Nexa",
                description=f"Olá **{interaction.user.global_name}**, \n \n <:setabrancaarrumada:1233574306294927430> Por favor, ao fazer uma pergunta, forneça uma **descrição completa** e **detalhada** do seu problema ou dúvida. \n \n <:setabrancaarrumada:1233574306294927430> **Evite mensagens curtas** como 'Oi' ou 'Olá', pois essas podem não fornecer informações suficientes para uma resposta adequada.",
                color=0x000000,
                url="https://agencynexa.com.br"
            )
            mensagem_embed.set_footer(text="Agency Nexa • © Todos os direitos reservados.", icon_url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
            mensagem_embed.set_thumbnail(url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")

            assumir_ticket = discord.ui.Button(
                style=discord.ButtonStyle.blurple,
                label="Assumir Ticket",
                emoji="<:suporte:1233336866489630730>",
                custom_id="assumir_ticket"
            )

            async def assumindoticket(interaction: discord.Interaction):
                if discord.utils.get(interaction.user.roles, id=1095021946850250834):
                    await interaction.response.send_message("<:tick:1233524716728352830> **Sucesso:** você assumiu o ticket.", ephemeral=True)
                    await interaction.channel.send(f"Seu ticket será atendido por {interaction.user.mention}.")
                    await interaction.channel.set_permissions(interaction.user, send_messages=True)

                    id_usuario = interaction.user.id
                    linha_encontrada = None
                    coluna_encontrada = None

                # Percorre todas as células da planilha
                for i, row in enumerate(sheet3.get_all_values(), start=1):
                    for j, cell in enumerate(row, start=1):
                        # Verifica se o valor da célula é igual ao ID do usuário
                        if cell == f"{id_usuario}":
                            # Se encontrou o ID do usuário na planilha, armazena a linha e a coluna e sai dos loops
                            linha_encontrada = i
                            coluna_encontrada = j
                            valor_atual = sheet3.cell(i, 3).value
                            print(valor_atual)
                            try:
                                valor_atual = int(valor_atual)
                            except ValueError:
                                print("Erro ao converter para valor inteiro")
                            novo_valor_total = valor_atual + 1
                            sheet3.update_cell(i, 3, str(novo_valor_total))
                            print("Número de tickets assumidos atualizado.")
                            break
                    if linha_encontrada is not None:
                        break

                    # Verifica se o usuário foi encontrado na planilha
                    if linha_encontrada is not None:
                        print(f"O usuário foi encontrado na linha {linha_encontrada} e coluna {coluna_encontrada}.")
                    else:
                        for row_index, row in enumerate(sheet3.get_all_values(), start=1):
                            if "linha vazia" in row:
                                sheet3.update_cell(row_index, 1, str(id_usuario))
                                sheet3.update_cell(row_index, 2, interaction.user.name)
                                sheet3.update_cell(row_index, 3, "1")
                                print("Usuário adicionado à planilha.")
                                break
                else:
                    await interaction.response.send_message("Você não é da **equipe de suporte**, apenas pessoas da equipe podem `Assumir Tickets`.", ephemeral=True)

            assumir_ticket.callback = assumindoticket

            fechar_ticket = discord.ui.Button(
                style=discord.ButtonStyle.red,
                label="Fechar Ticket",
                emoji="<:lixeira:1233523785202667642>",
                custom_id="fechar_ticket"
            )
            view = discord.ui.View()
            view.add_item(fechar_ticket)
            view.add_item(assumir_ticket)

            async def fechandoticket(interaction: discord.Interaction):
                if discord.utils.get(interaction.user.roles, id=1095021946850250834):
                    
                    tz_info = "America/Sao_Paulo"
                    
                    transcript = await chat_exporter.export(
                        novo_canal,
                        tz_info=tz_info,
                    )

                    transcript_file = discord.File(
                                io.BytesIO(transcript.encode()),
                                filename=f"transcript-{novo_canal.name}.html",
                            )
                    
                    horario_atual = datetime.datetime.now()
                    intervalo = datetime.timedelta(seconds=6)
                    horario_futuro = horario_atual + intervalo
                    timestamp_futuro = horario_futuro.timestamp()
                    timestamp_arredondado = round(timestamp_futuro)
                    await interaction.response.send_message(f"<:tick:1233524716728352830> Entendido, este ticket será fechado <t:{timestamp_arredondado}:R>.")
                    fechado_por = interaction.user.mention
                    fechado_por_id = interaction.user.id
                    await asyncio.sleep(5)
                    await interaction.channel.delete()

                    # Remove o ID do usuário da planilha
                    for i, row in enumerate(sheet2.get_all_values(), start=1):
                        if str(row[0]) == str(criado_por_id):
                            sheet2.update_cell(i, 1, "")  # Limpa a célula com o ID do usuário
                            sheet2.update_cell(i, 2, "")
                            break  # Interrompe o loop após encontrar e atualizar a célula
                            
                    server_id = '1195527779651964978'
                    channel_id = '1232547539962691617'

                    server = bot.get_guild(int(server_id))
                    if server is None:
                        print(f"Servidor não encontrado com o ID {server_id}")
                        return

                    channel = server.get_channel(int(channel_id))
                    if channel is None:
                        print(f"Canal não encontrado com o ID {channel_id}")
                        return

                    log_ticket = discord.Embed(
                        title="Ticket de suporte fechado",
                        color=0x00000,
                        url="https://agencynexa.com.br"
                    )

                    hora_atual = datetime.datetime.now()
                    time_atual = hora_atual.timestamp()
                    time_arredondado = round(time_atual)

                    log_ticket.add_field(name="Criado por:", value=f"{criado_por} `{criado_por_id}`", inline=False)
                    log_ticket.add_field(name="Fechado por:", value=f"{fechado_por} `{fechado_por_id}`", inline=False)
                    log_ticket.add_field(name="<:info:1196645186315493436> Acredita que foi um engano? Seu problema não foi resolvido?", value=f"\n Clique no botão abaixo e crie um novo ticket e entre em contato consco novamente!", inline=False)
                    log_ticket.add_field(name="<:relogio:1233543048051556453> Fechado em:", value=f"<t:{time_arredondado}:F>", inline=False)
                    log_ticket.set_footer(text="Agency Nexa • © Todos os direitos reservados.", icon_url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
                    log_ticket.set_thumbnail(url="https://cdn.discordapp.com/attachments/1118592272905285694/1122534689102647429/Design_sem_nome_1.png")
                    button_url = discord.ui.Button(
                        style=discord.ButtonStyle.link,
                        label="Criar novo ticket",
                        url="https://discordapp.com/channels/1094399810158739466/1095014679404875918"
                    )

                    view = discord.ui.View()
                    view.add_item(button_url)

                    await channel.send(embed=log_ticket, view=view)
                    await channel.send(file=transcript_file)

                    enviar = bot.get_user(criado_por_id)
                    await enviar.send(embed=log_ticket, view=view)
                    await enviar.send(file=transcript_file)

                else:
                    await interaction.response.send_message("<:x_:1233525194963030017>  **Erro:** Você não é da **equipe de suporte**, apenas pessoas da equipe podem `Fechar Tickets`.", ephemeral=True)

            fechar_ticket.callback = fechandoticket

            equipe = guild.get_role(1095021946850250834)

            content = f"{interaction.user.mention} {equipe.mention}"

            # Envio da mensagem embed no canal recém-criado
            mensagem_final = await novo_canal.send(embed=mensagem_embed, content=content, view=view)

            await mensagem_final.pin()

    reembolso_button.callback = partial(criar_Canal, channel_type=discord.ChannelType.text)

    botao = discord.ui.Button(
        style=discord.ButtonStyle.link,
        label="Visitar Agencynexa",
        url="https://agencynexa.com.br",
        emoji="<:nexa:1124517339354902638>"
    )
    view.add_item(botao)

    await interaction.channel.send(embed=embed, view=view)

bot.run(token)