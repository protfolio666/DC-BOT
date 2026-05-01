const {
  Client,
  GatewayIntentBits,
  PermissionsBitField,
  ChannelType,
  ActionRowBuilder,
  ButtonBuilder,
  ButtonStyle
} = require('discord.js');

const XLSX = require('xlsx');
const axios = require('axios');
const fs = require('fs');

const client = new Client({
  intents: [
    GatewayIntentBits.Guilds,
    GatewayIntentBits.GuildMessages,
    GatewayIntentBits.MessageContent
  ]
});

require('dotenv').config();

const TOKEN = process.env.TOKEN;
const GUILD_ID = process.env.GUILD_ID;
const CATEGORY_ID = process.env.CATEGORY_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const EXTRA_ROLE_ID = process.env.EXTRA_ROLE_ID;
const UPLOAD_CHANNEL_ID = process.env.UPLOAD_CHANNEL_ID;

client.once('clientReady', () => {
  console.log("Bot online");
});

// ---------- HELPERS ----------
const cleanId = (value) => value?.toString().replace(/[^0-9]/g, '').trim();

const resolveUser = async (guild, input) => {
  if (!input) return null;
  const id = cleanId(input);
  if (id.length >= 17) {
    try { return await guild.members.fetch(id); } catch {}
  }
  return null;
};

const resolveMultipleUsers = async (guild, input) => {
  if (!input) return [];
  const parts = input.toString().split(',');
  const users = [];

  for (let part of parts) {
    const user = await resolveUser(guild, part);
    if (user) users.push(user);
  }

  return users;
};

// ---------- EXCEL UPLOAD ----------
client.on('messageCreate', async message => {
  if (message.author.bot) return;
  if (message.channel.id !== UPLOAD_CHANNEL_ID) return;

  if (!message.member.permissions.has(PermissionsBitField.Flags.Administrator)) {
    return message.reply("❌ Only admins can upload matches.");
  }

  if (message.attachments.size === 0) return;

  const attachment = message.attachments.first();

  if (!attachment.name.endsWith('.xlsx')) {
    return message.reply("❌ Upload a .xlsx file");
  }

  await message.reply("📥 Processing Excel...");

  try {
    const filePath = `matches_${Date.now()}.xlsx`;

    const response = await axios.get(attachment.url, { responseType: 'arraybuffer' });
    fs.writeFileSync(filePath, response.data);

    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(sheet);

    const guild = await client.guilds.fetch(GUILD_ID);

    let matchNumber = 1;
    let errorLogs = [];

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const excelRow = i + 2;

      const round = row["Round"] || 1;

      const team1Name = row["Team Name 1"] || "Team A";
      const team2Name = row["Team Name 2"] || "Team B";

      const notes1 = row["Notes"] || "No extra info provided";
      const notes2 = row["Extra Notes"] || "";

      const extraTagsRaw = row["Extra Tags"] || "";

      const user1 = await resolveUser(guild, row["Player1 ID"]);
      const user2 = await resolveUser(guild, row["Player2 ID"]);

      const team1Users = await resolveMultipleUsers(guild, row["Player1 IDs"]);
      const team2Users = await resolveMultipleUsers(guild, row["Player2 IDs"]);

      if (team1Users.length === 0 && user1) team1Users.push(user1);
      if (team2Users.length === 0 && user2) team2Users.push(user2);

      if (team1Users.length === 0 || team2Users.length === 0) {
        errorLogs.push(`❌ Row ${excelRow} → Users not resolved`);
        continue;
      }

      try {
        const channel = await guild.channels.create({
          name: `round-${round}-match-${matchNumber}`,
          type: ChannelType.GuildText,
          parent: CATEGORY_ID,

          permissionOverwrites: [
            { id: guild.roles.everyone.id, deny: [PermissionsBitField.Flags.ViewChannel] },
            ...team1Users.map(u => ({
              id: u.id,
              allow: [PermissionsBitField.Flags.ViewChannel, PermissionsBitField.Flags.SendMessages]
            })),
            ...team2Users.map(u => ({
              id: u.id,
              allow: [PermissionsBitField.Flags.ViewChannel, PermissionsBitField.Flags.SendMessages]
            })),
            { id: EXTRA_ROLE_ID, allow: [PermissionsBitField.Flags.ViewChannel] }
          ]
        });

        const team1Mentions = team1Users.map(u => `<@${u.id}>`).join(' ');
        const team2Mentions = team2Users.map(u => `<@${u.id}>`).join(' ');

        const extraTags = extraTagsRaw
          .toString()
          .split(',')
          .filter(x => x.trim() !== "")
          .map(id => `<@${cleanId(id)}>`).join(' ');

        const rowBtn = new ActionRowBuilder().addComponents(
          new ButtonBuilder()
            .setCustomId(`close_${channel.id}`)
            .setLabel('Close Match')
            .setStyle(ButtonStyle.Danger)
        );

        await channel.send({
          content: `🏆 **Round ${round} - Match ${matchNumber}**
**${team1Name} vs ${team2Name}**

👥 ${team1Mentions} vs ${team2Mentions}

📌 ${notes1}
📌 ${notes2}

🔔 <@&${EXTRA_ROLE_ID}>
${extraTags}`,
          components: [rowBtn]
        });

      } catch (err) {
        errorLogs.push(`❌ Row ${excelRow} → ${err.message}`);
      }

      matchNumber++;
      await new Promise(r => setTimeout(r, 800));
    }

    fs.unlinkSync(filePath);

    if (errorLogs.length > 0) {
      await message.reply(`⚠️ Matches created with errors:\n\n${errorLogs.join('\n').slice(0, 1900)}`);
    } else {
      message.reply("✅ All matches created successfully!");
    }

  } catch (err) {
    console.log(err);
    message.reply("❌ Error processing file.");
  }
});

// ---------- CLOSE + TRANSCRIPT (UNCHANGED) ----------
client.on('interactionCreate', async interaction => {
  if (!interaction.isButton()) return;

  if (interaction.customId.startsWith('close_')) {

    if (
      !interaction.member.permissions.has(PermissionsBitField.Flags.Administrator) &&
      !interaction.member.roles.cache.has(EXTRA_ROLE_ID)
    ) {
      return interaction.reply({ content: "❌ Not allowed", flags: 64 });
    }

    const confirmRow = new ActionRowBuilder().addComponents(
      new ButtonBuilder().setCustomId(`confirm_${interaction.channel.id}`).setLabel('Yes').setStyle(ButtonStyle.Success),
      new ButtonBuilder().setCustomId(`cancel_${interaction.channel.id}`).setLabel('No').setStyle(ButtonStyle.Secondary)
    );

    return interaction.reply({
      content: "⚠️ Are you sure you want to close this match?",
      components: [confirmRow],
      flags: 64
    });
  }

  if (interaction.customId.startsWith('cancel_')) {
    return interaction.update({ content: "❌ Cancelled.", components: [] });
  }

  if (interaction.customId.startsWith('confirm_')) {

    const channel = interaction.channel;

    await interaction.update({ content: "Closing match...", components: [] });

    let allMessages = [];
    let lastId;

    while (true) {
      const options = { limit: 100 };
      if (lastId) options.before = lastId;

      const messages = await channel.messages.fetch(options);
      if (messages.size === 0) break;

      allMessages.push(...messages.values());
      lastId = messages.last().id;
    }

    const closedBy = interaction.user;

    const content = allMessages
      .map(m => {
        const name = m.author.discriminator !== "0"
          ? `${m.author.username}#${m.author.discriminator}`
          : m.author.username;

        const time = new Date(m.createdTimestamp).toLocaleString();

        const attachments = m.attachments.size > 0
          ? ` [Attachment: ${m.attachments.map(a => a.url).join(', ')}]`
          : '';

        return `[${time}] ${name}: ${m.content || "[No text]"}${attachments}`;
      })
      .reverse()
      .join('\n');

    const finalTranscript = `Closed By: ${closedBy.tag}\n\n${content}`;

    const buffer = Buffer.from(finalTranscript, 'utf-8');

    const logChannel = await interaction.guild.channels.fetch(LOG_CHANNEL_ID);

    await logChannel.send({
      content: `Transcript from ${channel.name}`,
      files: [{ attachment: buffer, name: `${channel.name}.txt` }]
    });

    setTimeout(async () => {
      await channel.delete();
    }, 3000);
  }
});

client.login(TOKEN);