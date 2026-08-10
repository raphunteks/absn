"use client";

import React from 'react';
import { AppWindow, LayoutTemplate, Database, Palette, Instagram, ChevronLeft } from 'lucide-react';

const teamData = [
  {
    category: "FullStack Development",
    icon: AppWindow,
    members: [
      { 
         name: "M. Aksa Arsyad, drg., S.KG", 
         role: "Lead Developer", 
         dept: "Dept. RKG RSIGM UMI", 
         img: "/axaprofil.jpg", 
         ig: "https://www.instagram.com/axaaxyz_01" 
      },
      { 
         name: "Abdullah H.D Lasari, drg M.Kom", 
         role: "Backend Development", 
         dept: "Dept. RKG RSIGM UMI", 
         img: "/public/dulprofil.png", 
         ig: "" 
      }
    ]
  },
  {
    category: "Frontend Development",
    icon: LayoutTemplate,
    members: [
      { 
         name: "M. Rakhmat Ersyad M, drg., Sp.RKG., S.H", 
         role: "Frontend Development", 
         dept: "Dept. RKG RSIGM UMI", 
         img: "/public/mrakhmatprofil.png", 
         ig: "https://www.instagram.com/mrakhmat" 
      },
      { 
         name: "Andi Nurul Azizah Tenrilili, drg., Sp.RKG", 
         role: "Frontend Development", 
         dept: "Dept. RKG RSIGM UMI", 
         img: "/public/azizahprofil.png", 
         ig: "" 
      }
    ]
  },
  {
    category: "Backend Development",
    icon: Database,
    members: [
      { 
         name: "Dian Handayani, drg., Sp.RKG., M.Kes", 
         role: "Backend Development", 
         dept: "Dept. RKG RSIGM UMI", 
         img: "/public/dianprofil.png", 
         ig: "" 
      }
    ]
  },
  {
    category: "UI/UX Design (CSS)",
    icon: Palette,
    members: [
      { 
         name: "Andi Rasdianti Inra Purnama, drg., Sp.RKG", 
         role: "UI/UX Design (CSS)", 
         dept: "Dept. RKG RSIGM UMI", 
         img: "/public/rasdiantiprofil.png", 
         ig: "" 
      }
    ]
  }
];

export default function OurTeam() {
  return (
    <div className="min-h-screen bg-[#f8fafc] flex flex-col items-center py-12 md:py-16 px-4 relative overflow-hidden font-sans">
      
      {/* Light Grid Background (Sesuai Referensi UI/UX) */}
      <div className="absolute inset-0 pointer-events-none bg-[linear-gradient(to_right,#80808015_1px,transparent_1px),linear-gradient(to_bottom,#80808015_1px,transparent_1px)] bg-[size:24px_24px]"></div>

      {/* Navigation Button */}
      <div className="w-full max-w-5xl mb-8 relative z-20 flex justify-start animate-in fade-in slide-in-from-top-5 duration-500">
        <a 
          href="/" 
          className="flex items-center gap-2 px-5 py-2.5 bg-white border border-slate-200 text-slate-600 font-bold text-sm rounded-full shadow-sm hover:shadow-md hover:border-[#0066cc] hover:text-[#0066cc] transition-all duration-300 active:scale-95"
        >
          <ChevronLeft className="w-4 h-4" /> Kembali ke Portal
        </a>
      </div>

      {/* Header Section */}
      <div className="text-center mb-16 relative z-10 animate-in fade-in slide-in-from-bottom-5 duration-700">
        <h1 className="text-4xl md:text-5xl font-black text-slate-800 tracking-tight">
          Developer <span className="text-[#0066cc]">Team</span>
        </h1>
        <p className="text-slate-500 mt-4 text-sm md:text-base font-medium max-w-lg mx-auto leading-relaxed">
          Sosok dibalik layar pengembangan Sistem Informasi BEM KBMFKG UMI & Dept. RKG RSIGM UMI.
        </p>
      </div>

      {/* Categories & Cards */}
      <div className="w-full max-w-5xl flex flex-col gap-16 relative z-10">
        {teamData.map((group, groupIdx) => (
          <div key={groupIdx} className="flex flex-col items-center animate-in fade-in slide-in-from-bottom-8 duration-700" style={{ animationDelay: `${groupIdx * 150}ms`, animationFillMode: 'both' }}>
            
            {/* Category Title Pill */}
            <div className="flex items-center gap-3 mb-8">
              <div className="w-12 h-10 bg-white border border-[#0066cc]/20 rounded-xl flex items-center justify-center shadow-sm text-[#0066cc]">
                <group.icon className="w-5 h-5" />
              </div>
              <h2 className="text-2xl md:text-3xl font-bold text-slate-800 border-b-4 border-[#0066cc] pb-1.5 tracking-tight">
                {group.category}
              </h2>
            </div>

            {/* Members Grid */}
            <div className="grid grid-cols-1 md:grid-cols-2 gap-6 w-full justify-center">
              {group.members.map((member, memberIdx) => (
                <div 
                  key={memberIdx} 
                  className={`bg-white border border-slate-100 p-8 rounded-[2rem] flex flex-col items-center text-center shadow-[0_10px_40px_rgba(0,0,0,0.04)] hover:shadow-[0_15px_50px_rgba(0,102,204,0.12)] transition-all duration-300 hover:-translate-y-1.5 ${group.members.length === 1 ? 'md:col-span-2 max-w-md mx-auto w-full' : 'w-full'}`}
                >
                  {/* Profile Image with Concentric Rings */}
                  <div className="w-28 h-28 md:w-32 md:h-32 bg-slate-50 rounded-full flex items-center justify-center border-[4px] border-slate-800 p-1 mb-5 relative shadow-lg">
                    {/* Outer Blue Ring Accent */}
                    <div className="absolute inset-0 rounded-full border-[3px] border-[#0066cc] scale-110 pointer-events-none opacity-80"></div>
                    {/* Glow effect behind image */}
                    <div className="absolute inset-0 bg-[#0066cc] blur-2xl opacity-20 rounded-full"></div>
                    
                    <div className="w-full h-full rounded-full overflow-hidden bg-slate-200 relative z-10">
                      <img 
                        src={member.img} 
                        alt={member.name} 
                        className="w-full h-full object-cover" 
                        onError={(e) => {
                          e.currentTarget.src = 'https://ui-avatars.com/api/?name=' + encodeURIComponent(member.name) + '&background=e2e8f0&color=0f172a';
                        }}
                      />
                    </div>
                  </div>

                  {/* Profile Info */}
                  <h3 className="text-lg md:text-xl font-black text-slate-800 tracking-tight leading-tight">
                    {member.name}
                  </h3>
                  <p className="text-[#0066cc] font-bold text-xs md:text-sm mt-1.5">
                    {member.role}
                  </p>
                  <p className="text-slate-400 text-[10px] md:text-xs font-medium mt-1 mb-6 max-w-[200px]">
                    {member.dept}
                  </p>

                  <div className="w-full h-[1px] bg-slate-100 mb-6"></div>

                  {/* Social Media Link */}
                  {member.ig ? (
                     <a 
                       href={member.ig} 
                       target="_blank" 
                       rel="noreferrer" 
                       className="w-10 h-10 bg-slate-50 hover:bg-[#0066cc] hover:text-white border border-slate-200 rounded-full flex items-center justify-center text-slate-500 transition-all duration-300 shadow-sm cursor-pointer active:scale-95 group"
                     >
                        <Instagram className="w-4 h-4 group-hover:scale-110 transition-transform" />
                     </a>
                  ) : (
                     <div className="w-10 h-10 bg-slate-50 border border-slate-100 rounded-full flex items-center justify-center text-slate-300 shadow-sm cursor-not-allowed">
                        <Instagram className="w-4 h-4" />
                     </div>
                  )}
                </div>
              ))}
            </div>

          </div>
        ))}
      </div>

    </div>
  );
}
