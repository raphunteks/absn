import { NextResponse } from 'next/server';

class Redis {
  url: string;
  token: string;
  constructor(config: { url: string; token: string }) {
    this.url = config.url || '';
    this.token = config.token || '';
    if (this.url.endsWith('/')) this.url = this.url.slice(0, -1);
  }
  static fromEnv() {
    return new Redis({ 
      url: process.env.NEXT_PUBLIC_UPSTASH_REDIS_REST_URL || process.env.NEXT_PUBLIC_KV_REST_API_URL || '', 
      token: process.env.NEXT_PUBLIC_UPSTASH_REDIS_REST_TOKEN || process.env.NEXT_PUBLIC_KV_REST_API_TOKEN || '' 
    });
  }
  async get(key: string) {
    try {
      const res = await fetch(this.url, { method: 'POST', headers: { Authorization: `Bearer ${this.token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(["GET", key]), cache: 'no-store' });
      const data = await res.json();
      return typeof data.result === 'string' ? JSON.parse(data.result) : data.result;
    } catch { return null; }
  }
  async set(key: string, value: any) {
    const strVal = typeof value === 'string' ? value : JSON.stringify(value);
    await fetch(this.url, { method: 'POST', headers: { Authorization: `Bearer ${this.token}`, 'Content-Type': 'application/json' }, body: JSON.stringify(["SET", key, strVal]) });
  }
}

// 1. BOT MENGAMBIL ANTRIAN PESAN (GET)
export async function GET() {
  const redis = Redis.fromEnv();
  const queue = await redis.get('axaxyz_wa_queue') || [];
  return NextResponse.json({ success: true, queue }, { status: 200 });
}

// 2. WEB MENAMBAH PESAN KE ANTRIAN (POST)
export async function POST(request: Request) {
  try {
    const { no_hp, formatted_message } = await request.json();
    if (!no_hp || !formatted_message) return NextResponse.json({ success: false, error: "no_hp dan formatted_message wajib diisi." }, { status: 400 });

    const redis = Redis.fromEnv();
    const currentQueue = await redis.get('axaxyz_wa_queue') || [];
    
    // Push pesan baru dengan ID Unik
    const newMessage = { id: Math.random().toString(36).substr(2, 9), target_number: no_hp, formatted_message };
    currentQueue.push(newMessage);
    
    await redis.set('axaxyz_wa_queue', currentQueue);

    return NextResponse.json({ success: true, message: "Pesan berhasil dimasukkan ke antrian (Queue).", data: newMessage }, { status: 200 });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}

// 3. BOT MENGHAPUS PESAN SETELAH TERKIRIM (DELETE)
export async function DELETE(request: Request) {
  try {
    const { message_id } = await request.json();
    if (!message_id) return NextResponse.json({ success: false, error: "message_id wajib dikirim" }, { status: 400 });

    const redis = Redis.fromEnv();
    let currentQueue = await redis.get('axaxyz_wa_queue') || [];
    
    // Filter out pesan yang id-nya sama dengan yang mau dihapus
    currentQueue = currentQueue.filter((msg: any) => msg.id !== message_id);
    
    await redis.set('axaxyz_wa_queue', currentQueue);

    return NextResponse.json({ success: true, message: "Pesan berhasil dihapus dari antrian." }, { status: 200 });
  } catch (error: any) {
    return NextResponse.json({ success: false, error: error.message }, { status: 500 });
  }
}
