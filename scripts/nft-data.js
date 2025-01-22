const { ethers } = require("ethers");
const { MongoClient } = require("mongodb");
const ExcelJS = require("exceljs");

// Ethereum provider setup
const ether_url = "https://eth-mainnet.g.alchemy.com/v2/nCgU9Tb7DOgNTu0c1epIqXRpgVyqyzAm"; // Alchemy RPC URL
const provider = new ethers.JsonRpcProvider(ether_url);

// Contract setup
const CONTRACT_ADDRESS = "0x7229600B699DD90B8e7bE0575d02DB58F409d2cB"; // Deployed contract address
const ABI = require("../utils/ethles.json"); // Contract ABI
const contract = new ethers.Contract(CONTRACT_ADDRESS, ABI, provider);

// MongoDB setup
const MONGO_URI = "mongodb://localhost:27017"; // MongoDB URI
const DB_NAME = "ethjets"; // Database name
const COLLECTION_NAME = "nft_balances"; // Collection name

// Block range
const START_BLOCK = 15118180; // First transaction block
const END_BLOCK = 21541992; // Last transaction block
const CHUNK_SIZE = 10000; // Number of blocks to process per batch

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
