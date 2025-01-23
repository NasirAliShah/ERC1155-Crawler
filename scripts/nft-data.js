const { ethers } = require("ethers");
const { MongoClient } = require("mongodb");
const ExcelJS = require("exceljs");
require("dotenv").config();

// Ethereum provider setup
const RPC_URL = process.env.RPC_URL || "";
const provider = new ethers.JsonRpcProvider(RPC_URL);

// Contract setup
const CONTRACT_ADDRESS = process.env.CONTRACT_ADDRESS || "";
const ABI = require("../utils/ethles.json"); // Contract ABI
const contract = new ethers.Contract(CONTRACT_ADDRESS, ABI, provider);

// MongoDB setup
const MONGO_URI = process.env.MONGO_URI || "";
const DB_NAME = process.env.DB_NAME || "";
const COLLECTION_NAME = process.env.COLLECTION_NAME || "";

// Block range
const START_BLOCK = process.env.START_BLOCK || 0; // First transaction block
const END_BLOCK = process.env.END_BLOCK || 0; // Last transaction block
const CHUNK_SIZE = process.env.CHUNK_SIZE || 0; // Number of blocks to process per batch

async function main() {
  // Connect to MongoDB
  const client = new MongoClient(MONGO_URI);
  await client.connect();
  const db = client.db(DB_NAME);
  const collection = db.collection(COLLECTION_NAME);

  // Clear MongoDB collection before processing
  await collection.deleteMany({});
  console.log("Cleared the nft_balances collection.");

  const blockCache = {}; // Cache block timestamps to reduce provider calls

  console.log(`Fetching events from block ${START_BLOCK} to ${END_BLOCK} in chunks of ${CHUNK_SIZE}...`);

  // Process blocks in chunks
  for (let fromBlock = START_BLOCK; fromBlock <= END_BLOCK; fromBlock += CHUNK_SIZE) {
    const toBlock = Math.min(fromBlock + CHUNK_SIZE - 1, END_BLOCK);

    try {
      // Fetch TransferSingle events
      const transferSingleFilter = contract.filters.TransferSingle();
      const events = await contract.queryFilter(transferSingleFilter, fromBlock, toBlock);
      console.log(`Fetched ${events.length} events from block ${fromBlock} to ${toBlock}.`);

      for (const event of events) {
        const { from, to, id, value } = event.args;

        // Fetch block timestamp and cache it
        if (!blockCache[event.blockNumber]) {
          const block = await provider.getBlock(event.blockNumber);
          blockCache[event.blockNumber] = block.timestamp;
        }
        const timestamp = blockCache[event.blockNumber];

        // Update MongoDB with the data
        if (to !== ethers.ZeroAddress) {
          await collection.updateOne(
            { address: to, tokenId: id.toString() }, // Match by address and token ID
            {
              $inc: { quantity: parseInt(value.toString()) }, // Increment the quantity
              $set: { lastUpdated: new Date(timestamp * 1000) }, // Set the timestamp
            },
            { upsert: true } // Insert if not exists
          );
        }

        // Handle burn transactions
        if (from !== ethers.ZeroAddress) {
          await collection.updateOne(
            { address: from, tokenId: id.toString() }, // Match by address and token ID
            {
              $inc: { quantity: -parseInt(value.toString()) }, // Decrement the quantity
              $set: { lastUpdated: new Date(timestamp * 1000) }, // Set the timestamp
            },
            { upsert: true } // Insert if not exists
          );
        }
      }
    } catch (err) {
      console.error(`Error fetching events from block ${fromBlock} to ${toBlock}:`, err);
    }
  }

  console.log("All events processed. Aggregating data...");

  // Aggregate data and export to Excel
  const allBalances = await collection.find({}).toArray();
  const aggregatedData = {};

  // Aggregate balances by address
  allBalances.forEach((balance) => {
    const { address, tokenId, quantity } = balance;

    if (!aggregatedData[address]) {
      aggregatedData[address] = { globalNFT: 0, captainNFT: 0, topGunNFT: 0 }; // Default to 0 for all NFTs
    }

    // Map tokenId to the appropriate NFT column
    if (tokenId === "1") {
      aggregatedData[address].globalNFT = quantity;
    } else if (tokenId === "2") {
      aggregatedData[address].captainNFT = quantity;
    } else if (tokenId === "3") {
      aggregatedData[address].topGunNFT = quantity;
    }
  });

  console.log("Aggregated data:", aggregatedData);

  // Generate Excel file
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("NFT Details");
  sheet.columns = [
    { header: "Address", key: "address", width: 42 },
    { header: "Global NFT (Token ID 1)", key: "globalNFT", width: 20 },
    { header: "Captain NFT (Token ID 2)", key: "captainNFT", width: 20 },
    { header: "Top Gun NFT (Token ID 3)", key: "topGunNFT", width: 20 },
  ];

  Object.entries(aggregatedData).forEach(([address, balances]) => {
    sheet.addRow({
      address,
      globalNFT: balances.globalNFT,
      captainNFT: balances.captainNFT,
      topGunNFT: balances.topGunNFT,
    });
  });

  await workbook.xlsx.writeFile("NFT_Details.xlsx");
  console.log("Excel file generated: NFT_Details.xlsx");

  // Close MongoDB connection
  await client.close();
}

main().catch((err) => {
  console.error("Error:", err);
});
