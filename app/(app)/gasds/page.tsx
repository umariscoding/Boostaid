import { PageHeader, Card } from '@/components/ui/primitives'

export default function GasdsPage() {
  return (
    <>
      <PageHeader title="GASDS" subtitle="Gift Aid Small Donations Scheme" />
      <Card className="p-8 max-w-2xl">
        <div className="font-display text-xl font-semibold text-ink">This section is coming next</div>
        <p className="text-muted mt-2 leading-relaxed">
          GASDS lets a charity claim a 25% top-up on small cash and contactless donations of £30 or less, with no donor
          declarations needed. Tracking it here, from your receipt book, is the next piece to add once the requirements
          are confirmed.
        </p>
      </Card>
    </>
  )
}
