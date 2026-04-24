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
require('dotenv').config(); // add at top

const TOKEN = process.env.TOKEN;
const GUILD_ID = process.env.GUILD_ID;
const CATEGORY_ID = process.env.CATEGORY_ID;
const LOG_CHANNEL_ID = process.env.LOG_CHANNEL_ID;
const EXTRA_ROLE_ID = process.env.EXTRA_ROLE_ID;
const UPLOAD_CHANNEL_ID = process.env.UPLOAD_CHANNEL_ID;

client.once('clientReady', () => {
  console.log("Bot online");
});

// ---------- HELPER ----------
const cleanId = (value) => value?.toString().replace(/[^0-9]/g, '').trim();

// ---------- USER RESOLVER ----------
const resolveUser = async (guild, input) => {
  if (!input) return null;

  const raw = input.toString().trim();

  const id = cleanId(raw);
  if (id.length >= 17) {
    try { return await guild.members.fetch(id); } catch {}
  }

  const mention = raw.match(/^<@!?(\d+)>$/);
  if (mention) {
    try { return await guild.members.fetch(mention[1]); } catch {}
  }

  await guild.members.fetch(); // ✅ ensures cache

  const lower = raw.toLowerCase();

  let member = guild.members.cache.find(m =>
    m.user.username.toLowerCase() === lower ||
    m.user.globalName?.toLowerCase() === lower ||
    m.nickname?.toLowerCase() === lower
  );

  if (member) return member;

  const tagMatch = raw.match(/^(.+)#(\d{4})$/);
  if (tagMatch) {
    const [_, name, disc] = tagMatch;

    member = guild.members.cache.find(m =>
      m.user.username.toLowerCase() === name.toLowerCase() &&
      m.user.discriminator === disc
    );

    if (member) return member;
  }

  return null;
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

    for (const row of data) {

      const round = row["Round"] || 1;

      const user1 = await resolveUser(guild, row["Player1 ID"]);
      const user2 = await resolveUser(guild, row["Player2 ID"]);

      if (!user1 || !user2) {
        console.log("❌ User not resolved:", row);
        continue;
      }

      const channel = await guild.channels.create({
        name: `R${round}-Match-${matchNumber}`,
        type: ChannelType.GuildText,
        parent: CATEGORY_ID,

        permissionOverwrites: [
          { id: guild.roles.everyone.id, deny: [PermissionsBitField.Flags.ViewChannel] },
          { id: user1.id, allow: [PermissionsBitField.Flags.ViewChannel, PermissionsBitField.Flags.SendMessages] },
          { id: user2.id, allow: [PermissionsBitField.Flags.ViewChannel, PermissionsBitField.Flags.SendMessages] },
          { id: EXTRA_ROLE_ID, allow: [PermissionsBitField.Flags.ViewChannel] }
        ]
      });

      const rowBtn = new ActionRowBuilder().addComponents(
        new ButtonBuilder()
          .setCustomId(`close_${channel.id}`)
          .setLabel('Close Match')
          .setStyle(ButtonStyle.Danger)
      );

      await channel.send({
        content: `🏆 **Round ${round}.${matchNumber}**\nMatch: <@${user1.id}> vs <@${user2.id}>`,
        components: [rowBtn]
      });

      matchNumber++;
      await new Promise(r => setTimeout(r, 800));
    }

    fs.unlinkSync(filePath);
    message.reply("✅ Matches created!");

  } catch (err) {
    console.log(err);
    message.reply("❌ Error processing file.");
  }
});

// ---------- CLOSE + CONFIRM + TRANSCRIPT ----------
client.on('interactionCreate', async interaction => {
  if (!interaction.isButton()) return;

  // STEP 1: CLOSE CLICK
  if (interaction.customId.startsWith('close_')) {

    if (
      !interaction.member.permissions.has(PermissionsBitField.Flags.Administrator) &&
      !interaction.member.roles.cache.has(EXTRA_ROLE_ID)
    ) {
      return interaction.reply({
        content: "❌ Not allowed",
        flags: 64
      });
    }

    const confirmRow = new ActionRowBuilder().addComponents(
      new ButtonBuilder()
        .setCustomId(`confirm_${interaction.channel.id}`)
        .setLabel('Yes')
        .setStyle(ButtonStyle.Success),

      new ButtonBuilder()
        .setCustomId(`cancel_${interaction.channel.id}`)
        .setLabel('No')
        .setStyle(ButtonStyle.Secondary)
    );

    return interaction.reply({
      content: "⚠️ Are you sure you want to close this match?",
      components: [confirmRow],
      flags: 64
    });
  }

  // CANCEL
  if (interaction.customId.startsWith('cancel_')) {
    return interaction.update({
      content: "❌ Cancelled.",
      components: []
    });
  }

  // CONFIRM CLOSE
  if (interaction.customId.startsWith('confirm_')) {

    const channel = interaction.channel;

    await interaction.update({
      content: "Closing match...",
      components: []
    });

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

    const finalTranscript =
`Closed By: ${closedBy.tag}\n\n${content}`;

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